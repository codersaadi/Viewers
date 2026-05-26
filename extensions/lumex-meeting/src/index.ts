import { Enums, EVENTS, eventTarget, utilities } from '@cornerstonejs/core';
import { annotation, Enums as ToolsEnums } from '@cornerstonejs/tools';

const id = '@lumex/extension-meeting';

type BridgeMode = 'disconnected' | 'presenter' | 'synced' | 'local_inspect';

type ParentMessage = {
  type?: string;
  sequence?: number;
  pointer?: {
    viewportId?: string;
    x?: number;
    y?: number;
    name?: string;
    color?: string;
    timestamp?: number;
  };
  artifact?: {
    kind?: string;
    measurement?: Record<string, any>;
  };
  action?: 'upsert' | 'delete';
  annotationUID?: string;
  state?: {
    studyInstanceUID?: string;
    seriesInstanceUID?: string;
    activeViewportId?: string;
    viewportLayout?: string;
    layout?: {
      numRows?: number;
      numCols?: number;
      layoutType?: string;
    };
    isHangingProtocolLayout?: boolean;
    hangingProtocol?: {
      protocolId?: string;
      stageIndex?: number;
    };
    layoutViewportCount?: number;
    viewports?: Array<{
      viewportId?: string;
      displaySetInstanceUID?: string;
      studyInstanceUID?: string;
      seriesInstanceUID?: string;
      imageIndex?: number;
      numberOfSlices?: number;
      zoom?: number;
      pan?: { x: number; y: number };
      camera?: Record<string, unknown>;
      voiRange?: { lower: number; upper: number };
      windowLevel?: { window: number; level: number };
      rotation?: number;
      flipHorizontal?: boolean;
      flipVertical?: boolean;
      invert?: boolean;
      isPlaying?: boolean;
      frameRate?: number;
    }>;
  };
};

type Runtime = {
  servicesManager?: AppTypes.ServicesManager;
  commandsManager?: any;
};
type ViewPort = any
type CaptureResult =
  | { state: Record<string, unknown>; reason?: never }
  | { state?: never; reason: string };

const isBridgeEnabled = () => {
  if (typeof window === 'undefined') {
    return false;
  }

  const params = new URLSearchParams(window.location.search);
  return params.get('lumexMeeting') === '1' && params.get('lumexBridge') === 'v2';
};

const createBridge = (runtime: Runtime) => {
  let mode: BridgeMode = 'disconnected';
  let lastViewerState = '';
  let lastCaptureFailure = '';
  let lastCaptureAttemptAt: number | null = null;
  let lastCaptureResult: 'success' | 'failed' | 'not_attempted' = 'not_attempted';
  let lastCaptureError: string | null = null;
  let latestApplySequence = 0;
  let boundViewportIds = '';
  let viewportEventDisposers: Array<() => void> = [];
  let readOnlyOverlays: HTMLDivElement[] = [];
  let presenterPointerOverlay: HTMLDivElement | null = null;
  let pointerFadeTimer: number | null = null;
  let pointerLabelFadeTimer: number | null = null;
  let lastPointerSentAt = 0;
  let knownMeasurementIds = new Set<string>();
  let lastAppliedLayoutSignature = '';
  let lastCapturedLayoutSignature = '';
  let emitViewerStateTimer: number | null = null;
  let pendingRemoteViewerState: ParentMessage | null = null;
  let remoteViewerStateApplyInFlight = false;

  const baseCapabilities = {
    bridgeVersion: '2.0',
    ohifVersion: '3.12.0-beta.135',
    sourceLevelBridge: true,
    canCapturePointer: true,
    canApplyReadOnly: true,
    canSyncAnnotations: true,
    canSyncCine: true,
  };

  const post = (message: Record<string, unknown>) => {
    try {
      window.parent.postMessage(message, window.location.origin);
    } catch {
      // Parent may be gone while OHIF is unloading.
    }
  };

  const services = () => runtime.servicesManager?.services ?? ({} as Record<string, any>);

  const controlledMeasurementTools = new Set([
    'Length',
    'Angle',
    'CobbAngle',
    'Probe',
    'DragProbe',
    'Bidirectional',
    'ArrowAnnotate',
    'EllipticalROI',
    'CircleROI',
    'RectangleROI',
    'PlanarFreehandROI',
    'SplineROI',
    'LivewireContour',
    'CalibrationLine',
    'AdvancedMagnify',
  ]);

  const rawAnnotationOnlyTools = new Set([
    'ArrowAnnotate',
    'EllipticalROI',
    'CircleROI',
    'RectangleROI',
    'PlanarFreehandROI',
    'SplineROI',
    'LivewireContour',
    'CalibrationLine',
    'AdvancedMagnify',
  ]);

  const isReadOnlySyncedMode = () => mode === 'synced';

  const blockSyncedUserInput = (event: Event) => {
    if (!isReadOnlySyncedMode()) {
      return;
    }

    event.preventDefault();
    event.stopImmediatePropagation();
  };

  const removeReadOnlyOverlays = () => {
    readOnlyOverlays.forEach(overlay => overlay.remove());
    readOnlyOverlays = [];
  };

  const getViewportElements = () => {
    const { cornerstoneViewportService, viewportGridService } = services();
    const state = viewportGridService?.getState?.();
    const viewportIds = Array.from(state?.viewports?.keys?.() ?? []) as string[];

    return viewportIds
      .map(viewportId => cornerstoneViewportService?.getCornerstoneViewport?.(viewportId)?.element)
      .filter(Boolean) as HTMLElement[];
  };

  const createReadOnlyOverlay = (element: HTMLElement) => {
    const rect = element.getBoundingClientRect();
    const overlay = document.createElement('div');
    overlay.setAttribute('data-lumex-read-only-overlay', 'true');
    overlay.setAttribute('aria-hidden', 'true');
    overlay.style.position = 'fixed';
    overlay.style.left = `${rect.left}px`;
    overlay.style.top = `${rect.top}px`;
    overlay.style.width = `${rect.width}px`;
    overlay.style.height = `${rect.height}px`;
    overlay.style.zIndex = '2147483647';
    overlay.style.pointerEvents = 'auto';
    overlay.style.cursor = 'not-allowed';
    overlay.style.background = 'transparent';
    overlay.title = 'Following expert view.';
    overlay.addEventListener('contextmenu', blockSyncedUserInput, true);
    overlay.addEventListener('dblclick', blockSyncedUserInput, true);
    overlay.addEventListener('mousedown', blockSyncedUserInput, true);
    overlay.addEventListener('mousemove', blockSyncedUserInput, true);
    overlay.addEventListener('mouseup', blockSyncedUserInput, true);
    overlay.addEventListener('pointerdown', blockSyncedUserInput, true);
    overlay.addEventListener('pointermove', blockSyncedUserInput, true);
    overlay.addEventListener('pointerup', blockSyncedUserInput, true);
    overlay.addEventListener('touchstart', blockSyncedUserInput, true);
    overlay.addEventListener('touchmove', blockSyncedUserInput, true);
    overlay.addEventListener('wheel', blockSyncedUserInput, { capture: true, passive: false });

    return overlay;
  };

  const updateReadOnlyOverlays = () => {
    if (!isReadOnlySyncedMode() || typeof document === 'undefined') {
      removeReadOnlyOverlays();
      return;
    }

    if (!document.body) {
      window.setTimeout(updateReadOnlyOverlays, 100);
      return;
    }

    const viewportElements = getViewportElements();
    if (!viewportElements.length) {
      removeReadOnlyOverlays();
      return;
    }

    removeReadOnlyOverlays();
    readOnlyOverlays = viewportElements.map(createReadOnlyOverlay);
    readOnlyOverlays.forEach(overlay => document.body.appendChild(overlay));
  };

  const updateReadOnlyGuard = () => {
    if (isReadOnlySyncedMode()) {
      updateReadOnlyOverlays();
      return;
    }

    removeReadOnlyOverlays();
  };

  const removePresenterPointerOverlay = () => {
    presenterPointerOverlay?.remove();
    presenterPointerOverlay = null;
    if (pointerFadeTimer !== null) {
      window.clearTimeout(pointerFadeTimer);
      pointerFadeTimer = null;
    }
    if (pointerLabelFadeTimer !== null) {
      window.clearTimeout(pointerLabelFadeTimer);
      pointerLabelFadeTimer = null;
    }
  };

  const getPointerViewportElement = (viewportId?: string) => {
    const { cornerstoneViewportService } = services();
    const candidates = [viewportId, getActiveViewportId(), getFirstViewportId()].filter(Boolean) as string[];

    for (const candidate of candidates) {
      const element = cornerstoneViewportService?.getCornerstoneViewport?.(candidate)?.element;
      if (element) {
        return element as HTMLElement;
      }
    }

    return null;
  };

  const applyPresenterPointer = (pointer?: ParentMessage['pointer']) => {
    if (mode === 'presenter' || !pointer || typeof pointer.x !== 'number' || typeof pointer.y !== 'number') {
      removePresenterPointerOverlay();
      return;
    }

    if (mode === 'local_inspect') {
      removePresenterPointerOverlay();
      return;
    }

    const element = getPointerViewportElement(pointer.viewportId);
    if (!element) {
      removePresenterPointerOverlay();
      return;
    }

    const clampedX = Math.min(1, Math.max(0, pointer.x));
    const clampedY = Math.min(1, Math.max(0, pointer.y));
    const color = pointer.color || '#3b82f6';

    if (!presenterPointerOverlay) {
      const overlay = document.createElement('div');
      overlay.setAttribute('data-lumex-presenter-pointer', 'true');
      overlay.style.position = 'absolute';
      overlay.style.zIndex = '1000';
      overlay.style.pointerEvents = 'none';
      overlay.style.transform = 'translate(-50%, -50%)';
      overlay.style.transition = 'left 75ms linear, top 75ms linear, opacity 180ms ease';
      overlay.innerHTML = `
        <div data-lumex-pointer-ring style="position:absolute;width:10px;height:10px;left:-5px;top:-5px;border-radius:9999px;border:1px solid rgba(255,255,255,.75);box-shadow:0 0 8px rgba(34,211,238,.45);"></div>
        <div data-lumex-pointer-dot style="position:relative;width:6px;height:6px;left:-3px;top:-3px;border-radius:9999px;border:1px solid rgba(255,255,255,.85);opacity:.9;box-shadow:0 0 8px rgba(34,211,238,.45);"></div>
        <div data-lumex-pointer-label style="position:absolute;top:-18px;left:14px;white-space:nowrap;border-radius:3px;padding:1px 5px;font:500 9px/12px sans-serif;color:white;background:rgba(8,47,73,.62);box-shadow:0 2px 8px rgba(0,0,0,.22);opacity:0;transition:opacity 180ms ease;"></div>
      `;
      presenterPointerOverlay = overlay;
    }

    if (presenterPointerOverlay.parentElement !== element) {
      presenterPointerOverlay.remove();
      if (window.getComputedStyle(element).position === 'static') {
        element.style.position = 'relative';
      }
      element.appendChild(presenterPointerOverlay);
    }

    presenterPointerOverlay.style.left = `${clampedX * 100}%`;
    presenterPointerOverlay.style.top = `${clampedY * 100}%`;
    presenterPointerOverlay.style.opacity = '1';
    const ring = presenterPointerOverlay.querySelector<HTMLElement>('[data-lumex-pointer-ring]')!;
    const dot = presenterPointerOverlay.querySelector<HTMLElement>('[data-lumex-pointer-dot]')!;
    const label = presenterPointerOverlay.querySelector<HTMLElement>('[data-lumex-pointer-label]')!;
    ring.style.borderColor = color;
    dot.style.backgroundColor = color;
    label.textContent = pointer.name || 'Presenter';
    label.style.opacity = '1';

    if (pointerLabelFadeTimer !== null) {
      window.clearTimeout(pointerLabelFadeTimer);
    }
    pointerLabelFadeTimer = window.setTimeout(() => {
      if (presenterPointerOverlay) {
        presenterPointerOverlay.querySelector<HTMLElement>('[data-lumex-pointer-label]')!.style.opacity = '0';
      }
    }, 1000);

    if (pointerFadeTimer !== null) {
      window.clearTimeout(pointerFadeTimer);
    }
    pointerFadeTimer = window.setTimeout(() => {
      if (presenterPointerOverlay) {
        presenterPointerOverlay.style.opacity = '0';
      }
    }, 1600);
  };

  const getCanCaptureViewportState = () => {
    const { viewportGridService, cornerstoneViewportService, displaySetService } = services();
    return !!(
      runtime.servicesManager?.services &&
      viewportGridService?.getState &&
      viewportGridService?.getActiveViewportId &&
      cornerstoneViewportService?.getCornerstoneViewport &&
      displaySetService?.getDisplaySetByUID
    );
  };

  const getCanApplyViewportState = () => {
    const { viewportGridService, cornerstoneViewportService, displaySetService } = services();
    return !!(
      runtime.servicesManager?.services &&
      viewportGridService?.getState &&
      viewportGridService?.setDisplaySetsForViewport &&
      cornerstoneViewportService?.getCornerstoneViewport &&
      displaySetService?.getDisplaySetByUID &&
      displaySetService?.getDisplaySetsForSeries
    );
  };

  const getCapabilities = () => ({
    ...baseCapabilities,
    canCaptureViewportState: getCanCaptureViewportState(),
    canApplyViewportState: getCanApplyViewportState(),
  });

  const emitCaptureFailure = (reason: string) => {
    lastCaptureResult = 'failed';
    lastCaptureError = reason;

    if (reason === lastCaptureFailure) {
      return;
    }

    lastCaptureFailure = reason;
    post({ type: 'local_error', message: `capture_failed: ${reason}` });
  };

  const getCaptureDiagnostics = () => {
    const { viewportGridService, cornerstoneViewportService } = services();
    const state = viewportGridService?.getState?.();
    const activeViewportId = getActiveViewportId();
    const viewportIds = Array.from(state?.viewports?.keys?.() ?? []) as string[];
    const cornerstoneViewport = activeViewportId
      ? cornerstoneViewportService?.getCornerstoneViewport?.(activeViewportId)
      : null;

    return {
      servicesAvailable: !!runtime.servicesManager?.services,
      activeViewportId: activeViewportId ?? null,
      viewportCount: viewportIds.length,
      hasCornerstoneViewport: !!cornerstoneViewport,
      hasElement: !!cornerstoneViewport?.element,
      lastCaptureAttemptAt,
      lastCaptureResult,
      lastCaptureError,
    };
  };

  const getMeasurementId = (measurement: any) => {
    const uid = measurement?.uid ?? measurement?.annotationUID;
    return typeof uid === 'string' ? uid : undefined;
  };

  const getMeasurementFromEvent = (payload: any) => {
    if (!payload) {
      return undefined;
    }

    return payload.measurement
      ?? payload.annotation
      ?? payload.detail?.measurement
      ?? payload.detail?.annotation
      ?? payload;
  };

  const normalizeSerializedValue = (value: any): any => {
    if (!value || typeof value !== 'object') {
      return value;
    }

    if (Array.isArray(value)) {
      return value.map(normalizeSerializedValue);
    }

    const keys = Object.keys(value);
    const numericKeys = keys.filter(key => /^\d+$/.test(key));
    if (numericKeys.length === keys.length && numericKeys.length > 0) {
      return numericKeys
        .sort((a, b) => Number(a) - Number(b))
        .map(key => normalizeSerializedValue(value[key]));
    }

    return Object.fromEntries(keys.map(key => [key, normalizeSerializedValue(value[key])]));
  };

  const cloneSerializable = (value: any) => {
    if (!value || typeof value !== 'object') {
      return value;
    }

    try {
      return normalizeSerializedValue(JSON.parse(JSON.stringify(value)));
    } catch {
      return null;
    }
  };

  const isValidPoint = (point: any) => (
    Array.isArray(point)
    && point.length >= 2
    && point.every((value: any) => typeof value === 'number' && Number.isFinite(value))
  );

  const hasClosedContour = (value: any) => Array.isArray(value) && value.length >= 3 && value.every(isValidPoint);

  const getAnnotationContour = (value: any) => value?.data?.contour?.polyline;

  const hasFullRenderableAnnotation = (measurement: any) => {
    const annotationPayload = measurement?.annotation ?? measurement;
    return !!annotationPayload?.annotationUID && hasClosedContour(getAnnotationContour(annotationPayload));
  };

  const hasRenderableGeometry = (measurement: any) => {
    const data = measurement?.data ?? {};
    const points = data?.handles?.points ?? measurement?.points ?? measurement?.handles?.points;
    const contourPoints = data?.contour?.polyline;

    if (Array.isArray(points) && points.length > 0) {
      return points.every(isValidPoint);
    }

    if (Array.isArray(contourPoints)) {
      return hasClosedContour(contourPoints);
    }

    return false;
  };

  const isControlledMeasurement = (measurement: any) => {
    const toolName = measurement?.toolName ?? measurement?.metadata?.toolName;
    return typeof toolName === 'string' && controlledMeasurementTools.has(toolName);
  };

  const sanitizeMeasurement = (measurement: any) => {
    const data = measurement?.data ?? {};
    const metadata = measurement?.metadata ?? {};
    const cachedStats = data?.cachedStats ?? measurement?.cachedStats ?? data ?? null;
    const primaryDisplayText = Array.isArray(measurement?.displayText?.primary)
      ? measurement.displayText.primary[0]
      : null;
    const fullAnnotation = measurement?.annotationUID ? cloneSerializable(measurement) : null;

    return {
      uid: getMeasurementId(measurement) ?? null,
      toolName: measurement?.toolName ?? metadata?.toolName ?? null,
      metadata,
      data,
      annotation: fullAnnotation,
      label: measurement?.label ?? data?.label ?? null,
      description: measurement?.description ?? null,
      text: measurement?.text ?? data?.text ?? measurement?.label ?? primaryDisplayText ?? null,
      unit: measurement?.unit ?? null,
      displayText: measurement?.displayText ?? null,
      length: typeof measurement?.length === 'number' ? measurement.length : null,
      area: typeof measurement?.area === 'number' ? measurement.area : null,
      perimeter: typeof measurement?.perimeter === 'number' ? measurement.perimeter : null,
      mean: typeof measurement?.mean === 'number' ? measurement.mean : null,
      stdDev: typeof measurement?.stdDev === 'number' ? measurement.stdDev : null,
      shortestDiameter: typeof measurement?.shortestDiameter === 'number' ? measurement.shortestDiameter : null,
      longestDiameter: typeof measurement?.longestDiameter === 'number' ? measurement.longestDiameter : null,
      points: measurement?.points ?? null,
      textBox: measurement?.textBox ?? data?.textBox ?? null,
      handles: data?.handles ?? null,
      cachedStats,
      referenceStudyUID: measurement?.referenceStudyUID ?? metadata?.StudyInstanceUID ?? null,
      referenceSeriesUID: measurement?.referenceSeriesUID ?? metadata?.SeriesInstanceUID ?? null,
      SOPInstanceUID: measurement?.SOPInstanceUID ?? metadata?.SOPInstanceUID ?? null,
      FrameOfReferenceUID: measurement?.FrameOfReferenceUID ?? metadata?.FrameOfReferenceUID ?? null,
      frameNumber: measurement?.frameNumber ?? data?.frameNumber ?? null,
      displaySetInstanceUID: measurement?.displaySetInstanceUID ?? metadata?.displaySetInstanceUID ?? null,
      referencedImageId: measurement?.referencedImageId ?? metadata?.referencedImageId ?? null,
    };
  };

  const emitMeasurementChanged = (action: 'created' | 'updated', measurementEvent?: any) => {
    const measurement = getMeasurementFromEvent(measurementEvent);
    if (mode !== 'presenter' || !measurement || !isControlledMeasurement(measurement)) {
      return;
    }

    const uid = getMeasurementId(measurement);
    if (!uid) {
      return;
    }

    const toolName = measurement?.toolName ?? measurement?.metadata?.toolName;
    if (rawAnnotationOnlyTools.has(toolName) && !hasFullRenderableAnnotation(measurement)) {
      return;
    }

    if (!hasRenderableGeometry(measurement)) {
      return;
    }

    let eventAction = action;
    if (!knownMeasurementIds.has(uid)) {
      knownMeasurementIds.add(uid);
      eventAction = 'created';
    } else if (action === 'created') {
      return;
    }

    post({
      type: eventAction === 'created' ? 'annotation_created' : 'annotation_updated',
      annotation: {
        kind: 'measurement',
        capturedAt: new Date().toISOString(),
        measurement: sanitizeMeasurement(measurement),
      },
    });
  };

  const emitMeasurementCreated = (event: any) => emitMeasurementChanged('created', event);
  const emitMeasurementUpdated = (event: any) => emitMeasurementChanged('updated', event);

  const emitMeasurementRemoved = (event: any) => {
    if (mode !== 'presenter') {
      return;
    }

    const measurement = getMeasurementFromEvent(event);
    const uid = typeof measurement === 'string' ? measurement : getMeasurementId(measurement);
    if (!uid) {
      return;
    }

    knownMeasurementIds.delete(uid);
    post({
      type: 'annotation_deleted',
      annotationUID: uid,
    });
  };

  const seedKnownMeasurements = () => {
    const measurementService = services().measurementService;
    const measurements = measurementService?.getMeasurements?.() ?? [];
    knownMeasurementIds = new Set(
      measurements
        .map(getMeasurementId)
        .filter((uid : string | null): uid is string => typeof uid === 'string')
    );
  };

  const applyMeasurementArtifact = (message: ParentMessage) => {
    const action = message.action ?? 'upsert';
    if (mode !== 'synced') {
      return;
    }

    if (action === 'delete') {
      const annotationUID = message.annotationUID;
      if (!annotationUID) {
        post({ type: 'measurement_artifact_apply_failed', action, reason: 'annotation_uid_missing' });
        return;
      }

      try {
        annotation.state.removeAnnotation?.(annotationUID);
        services().cornerstoneViewportService?.getRenderingEngine?.()?.render?.();
        post({ type: 'measurement_artifact_applied', action, annotationUID });
      } catch (error) {
        post({
          type: 'measurement_artifact_apply_failed',
          action,
          annotationUID,
          reason: error instanceof Error ? error.message : 'Failed to remove measurement artifact',
        });
      }
      return;
    }

    const artifact = message.artifact;
    if (!artifact?.measurement) {
      post({ type: 'measurement_artifact_apply_failed', action, reason: 'measurement_missing' });
      return;
    }

    const measurement = artifact.measurement;
    const uid = typeof measurement.uid === 'string' ? measurement.uid : undefined;
    const toolName = typeof measurement.toolName === 'string' ? measurement.toolName : undefined;
    if (!uid || !toolName || !controlledMeasurementTools.has(toolName)) {
      post({ type: 'measurement_artifact_apply_failed', action, annotationUID: uid, reason: 'unsupported_measurement' });
      return;
    }

    try {
      if (rawAnnotationOnlyTools.has(toolName) && !measurement.annotation) {
        post({ type: 'measurement_artifact_apply_failed', action, annotationUID: uid, reason: 'raw_annotation_missing' });
        return;
      }

      if (measurement.annotation && typeof measurement.annotation === 'object') {
        const fullAnnotation = cloneSerializable(measurement.annotation);
        if (!fullAnnotation?.data || !fullAnnotation?.metadata) {
          post({ type: 'measurement_artifact_apply_failed', action, annotationUID: uid, reason: 'annotation_payload_invalid' });
          return;
        }
        if (rawAnnotationOnlyTools.has(toolName) && !hasClosedContour(fullAnnotation.data?.contour?.polyline)) {
          post({ type: 'measurement_artifact_apply_failed', action, annotationUID: uid, reason: 'contour_incomplete' });
          return;
        }

        if (annotation.state.getAnnotation?.(uid)) {
          annotation.state.removeAnnotation?.(uid);
        }

        fullAnnotation.annotationUID = uid;
        fullAnnotation.metadata = {
          ...fullAnnotation.metadata,
          toolName,
        };
        fullAnnotation.highlighted = false;
        fullAnnotation.isLocked = true;
        fullAnnotation.invalidated = true;
        annotation.state.addAnnotation(fullAnnotation);
        services().cornerstoneViewportService?.getRenderingEngine?.()?.render?.();
        post({ type: 'measurement_artifact_applied', action, annotationUID: uid });
        return;
      }

      const handles = measurement.handles && typeof measurement.handles === 'object'
        ? measurement.handles
        : measurement.points
          ? { points: measurement.points }
          : undefined;
      if (!handles) {
        post({ type: 'measurement_artifact_apply_failed', action, annotationUID: uid, reason: 'handles_missing' });
        return;
      }

      if (annotation.state.getAnnotation?.(uid)) {
        annotation.state.removeAnnotation?.(uid);
      }

      annotation.state.addAnnotation({
        annotationUID: uid,
        highlighted: false,
        isLocked: true,
        invalidated: true,
        metadata: {
          ...(measurement.metadata && typeof measurement.metadata === 'object' ? measurement.metadata : {}),
          toolName,
          FrameOfReferenceUID: measurement.FrameOfReferenceUID ?? undefined,
          referencedImageId: measurement.referencedImageId ?? undefined,
          StudyInstanceUID: measurement.referenceStudyUID ?? undefined,
          SeriesInstanceUID: measurement.referenceSeriesUID ?? undefined,
          SOPInstanceUID: measurement.SOPInstanceUID ?? undefined,
          displaySetInstanceUID: measurement.displaySetInstanceUID ?? undefined,
        },
        data: {
          ...(measurement.data && typeof measurement.data === 'object' ? measurement.data : {}),
          label: measurement.label ?? undefined,
          text: measurement.text ?? undefined,
          handles,
          cachedStats: measurement.cachedStats ?? {},
          textBox: measurement.textBox ?? undefined,
          frameNumber: measurement.frameNumber ?? undefined,
        },
      });

      services().cornerstoneViewportService?.getRenderingEngine?.()?.render?.();
      post({ type: 'measurement_artifact_applied', action, annotationUID: uid });
    } catch (error) {
      post({
        type: 'measurement_artifact_apply_failed',
        action,
        annotationUID: uid,
        reason: error instanceof Error ? error.message : 'Failed to apply measurement artifact',
      });
    }
  };

  const getActiveViewportId = () => {
    const { viewportGridService } = services();
    return viewportGridService?.getActiveViewportId?.() ?? viewportGridService?.getState?.()?.activeViewportId;
  };

  const getViewportLayout = () => {
    const layout = services().viewportGridService?.getState?.()?.layout;
    if (!layout) {
      return 'unknown';
    }

    return layout.layoutType || `${layout.numRows}x${layout.numCols}`;
  };

  const getLayoutState = () => {
    const { viewportGridService, hangingProtocolService } = services();
    const state = viewportGridService?.getState?.();
    const layout = state?.layout;
    const hangingProtocol = hangingProtocolService?.getState?.();

    return {
      layout: layout
        ? {
            numRows: layout.numRows,
            numCols: layout.numCols,
            layoutType: layout.layoutType,
          }
        : undefined,
      isHangingProtocolLayout: !!state?.isHangingProtocolLayout,
      hangingProtocol: hangingProtocol?.protocolId
        ? {
            protocolId: hangingProtocol.protocolId,
            stageIndex: typeof hangingProtocol.stageIndex === 'number' ? hangingProtocol.stageIndex : undefined,
          }
        : undefined,
    };
  };

  const getImageIndex = (viewport : ViewPort) => {
    try {
      if (typeof viewport?.getCurrentImageIdIndex === 'function') {
        return viewport.getCurrentImageIdIndex();
      }
      if (typeof viewport?.getSliceIndex === 'function') {
        return viewport.getSliceIndex();
      }
      if (typeof viewport?.currentImageIdIndex === 'number') {
        return viewport.currentImageIdIndex;
      }
      const sliceData = utilities.getImageSliceDataForVolumeViewport?.(viewport);
      if (typeof sliceData?.imageIndex === 'number') {
        return sliceData.imageIndex;
      }
    } catch {
      return 0;
    }

    return 0;
  };

  const getImageIds = (viewport : ViewPort) => {
    try {
      if (typeof viewport?.getImageIds === 'function') {
        return viewport.getImageIds() ?? [];
      }
      if (Array.isArray(viewport?.imageIds)) {
        return viewport.imageIds;
      }
      return utilities.getViewportImageIds?.(viewport) ?? [];
    } catch {
      return [];
    }

    return [];
  };

  const getCurrentImageId = (viewport : ViewPort, imageIds: string[], imageIndex: number) => {
    try {
      if (typeof viewport?.getCurrentImageId === 'function') {
        return viewport.getCurrentImageId();
      }
    } catch {
      return imageIds[imageIndex] ?? null;
    }

    return imageIds[imageIndex] ?? null;
  };

  const getNumberOfSlices = (viewport : ViewPort, imageIds: string[], displaySet : {
    numImageFrames?: number;
    instances?: { length?: number }[];
  }) => {
    try {
      const sliceData = utilities.getImageSliceDataForVolumeViewport?.(viewport);
      if (typeof sliceData?.numberOfSlices === 'number') {
        return sliceData.numberOfSlices;
      }
      if (typeof viewport?.getNumberOfSlices === 'function') {
        return viewport.getNumberOfSlices();
      }
    } catch {
      return imageIds.length || displaySet?.numImageFrames || displaySet?.instances?.length || 1;
    }

    return imageIds.length || displaySet?.numImageFrames || displaySet?.instances?.length || 1;
  };

  const getCameraState = (viewport : ViewPort) => {
    try {
      if (typeof viewport?.getCamera !== 'function') {
        return null;
      }

      const camera = viewport.getCamera();
      if (!camera) {
        return null;
      }

      return {
        parallelScale: camera.parallelScale,
        focalPoint: camera.focalPoint,
        position: camera.position,
        viewUp: camera.viewUp,
        viewPlaneNormal: camera.viewPlaneNormal,
        flipHorizontal: camera.flipHorizontal,
        flipVertical: camera.flipVertical,
      };
    } catch {
      return null;
    }
  };

  const getZoomPanState = (viewport : ViewPort) => {
    const camera = getCameraState(viewport);
    if (!camera) {
      return { camera: null, zoom: undefined, pan: undefined };
    }

    const parallelScale = typeof camera.parallelScale === 'number' ? camera.parallelScale : undefined;
    const focalPoint = Array.isArray(camera.focalPoint) ? camera.focalPoint : [];

    return {
      camera,
      zoom: parallelScale && parallelScale !== 0 ? 1 / parallelScale : undefined,
      pan: typeof focalPoint[0] === 'number' && typeof focalPoint[1] === 'number'
        ? { x: focalPoint[0], y: focalPoint[1] }
        : undefined,
    };
  };

  const getViewportVolumeId = (viewport : ViewPort, displaySetInstanceUID?: string) => {
    try {
      if (typeof viewport?.getAllVolumeIds !== 'function') {
        return undefined;
      }

      const volumeIds = viewport.getAllVolumeIds() ?? [];
      if (!displaySetInstanceUID) {
        return volumeIds[0];
      }

      return volumeIds.find((volumeId: string) => volumeId.includes(displaySetInstanceUID)) ?? volumeIds[0];
    } catch {
      return undefined;
    }
  };

  const getWindowLevelState = (viewport : ViewPort, displaySetInstanceUID?: string) => {
    try {
      if (typeof viewport?.getProperties !== 'function') {
        return { voiRange: undefined, windowLevel: undefined };
      }

      const volumeId = getViewportVolumeId(viewport, displaySetInstanceUID);
      const properties = volumeId ? viewport.getProperties(volumeId) : viewport.getProperties();
      const voiRange = properties?.voiRange;
      if (typeof voiRange?.lower !== 'number' || typeof voiRange?.upper !== 'number') {
        return { voiRange: undefined, windowLevel: undefined };
      }

      const window = voiRange.upper - voiRange.lower;
      return {
        voiRange: { lower: voiRange.lower, upper: voiRange.upper },
        windowLevel: {
          window,
          level: voiRange.lower + window / 2,
        },
      };
    } catch {
      return { voiRange: undefined, windowLevel: undefined };
    }
  };

  const getPresentationState = (viewport : ViewPort, displaySetInstanceUID?: string) => {
    try {
      const presentation = typeof viewport?.getViewPresentation === 'function'
        ? viewport.getViewPresentation()
        : {};
      const camera = typeof viewport?.getCamera === 'function' ? viewport.getCamera() : {};
      const volumeId = getViewportVolumeId(viewport, displaySetInstanceUID);
      const properties = typeof viewport?.getProperties === 'function'
        ? volumeId ? viewport.getProperties(volumeId) : viewport.getProperties()
        : {};

      return {
        rotation: typeof presentation?.rotation === 'number' ? presentation.rotation : undefined,
        flipHorizontal: typeof presentation?.flipHorizontal === 'boolean'
          ? presentation.flipHorizontal
          : typeof camera?.flipHorizontal === 'boolean'
            ? camera.flipHorizontal
            : undefined,
        flipVertical: typeof presentation?.flipVertical === 'boolean'
          ? presentation.flipVertical
          : typeof camera?.flipVertical === 'boolean'
            ? camera.flipVertical
            : undefined,
        invert: typeof properties?.invert === 'boolean' ? properties.invert : undefined,
      };
    } catch {
      return {
        rotation: undefined,
        flipHorizontal: undefined,
        flipVertical: undefined,
        invert: undefined,
      };
    }
  };

  const getCineState = (viewportId: string) => {
    try {
      const state = services().cineService?.getState?.();
      const cine = state?.cines?.[viewportId];
      return {
        isPlaying: !!cine?.isPlaying,
        frameRate: typeof cine?.frameRate === 'number' ? cine.frameRate : undefined,
      };
    } catch {
      return { isPlaying: false, frameRate: undefined };
    }
  };

  const applyCineState = (
    viewportId: string,
    element: HTMLElement,
    viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]
  ) => {
    const cineService = services().cineService;
    if (!cineService || typeof viewportState.isPlaying !== 'boolean') {
      return;
    }

    const frameRate = typeof viewportState.frameRate === 'number' ? Math.max(viewportState.frameRate, 1) : 24;
    try {
      cineService.setIsCineEnabled?.(viewportState.isPlaying);
      cineService.setCine?.({
        id: viewportId,
        isPlaying: viewportState.isPlaying,
        frameRate,
      });
      if (!viewportState.isPlaying) {
        cineService.stopClip?.(element, { viewportId });
      }
    } catch {
      // Cine controls are best-effort after the frame state has applied.
    }
  };

  const getTargetVOIRange = (viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]) => {
    const voiRange = viewportState.voiRange;
    if (typeof voiRange?.lower === 'number' && typeof voiRange?.upper === 'number') {
      return { lower: voiRange.lower, upper: voiRange.upper };
    }

    const windowLevel = viewportState.windowLevel;
    if (typeof windowLevel?.window === 'number' && typeof windowLevel?.level === 'number') {
      return {
        lower: windowLevel.level - windowLevel.window / 2,
        upper: windowLevel.level + windowLevel.window / 2,
      };
    }

    return null;
  };

  const areVOIRangesClose = (
    first?: { lower: number; upper: number } | null,
    second?: { lower: number; upper: number } | null
  ) => {
    if (!first || !second) {
      return false;
    }

    return Math.abs(first.lower - second.lower) < 0.01 && Math.abs(first.upper - second.upper) < 0.01;
  };

  const applyWindowLevelState = (
    viewport: any,
    viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]
  ) => {
    const voiRange = getTargetVOIRange(viewportState);
    if (!voiRange || typeof viewport?.setProperties !== 'function') {
      return false;
    }

    const volumeId = getViewportVolumeId(viewport, viewportState.displaySetInstanceUID);
    if (volumeId) {
      viewport.setProperties({ voiRange }, volumeId);
    } else {
      viewport.setProperties({ voiRange });
    }
    viewport.render?.();
    return true;
  };

  const applyPresentationState = (
    viewport: any,
    viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]
  ) => {
    const presentationPatch: Record<string, unknown> = {};
    if (typeof viewportState.rotation === 'number') {
      presentationPatch.rotation = ((viewportState.rotation % 360) + 360) % 360;
    }
    if (typeof viewportState.flipHorizontal === 'boolean') {
      presentationPatch.flipHorizontal = viewportState.flipHorizontal;
    }
    if (typeof viewportState.flipVertical === 'boolean') {
      presentationPatch.flipVertical = viewportState.flipVertical;
    }

    if (Object.keys(presentationPatch).length && typeof viewport?.setViewPresentation === 'function') {
      viewport.setViewPresentation(presentationPatch);
    } else if (typeof viewport?.setCamera === 'function') {
      const cameraPatch: Record<string, unknown> = {};
      if (typeof viewportState.flipHorizontal === 'boolean') {
        cameraPatch.flipHorizontal = viewportState.flipHorizontal;
      }
      if (typeof viewportState.flipVertical === 'boolean') {
        cameraPatch.flipVertical = viewportState.flipVertical;
      }
      if (Object.keys(cameraPatch).length) {
        viewport.setCamera(cameraPatch);
      }
    }

    if (typeof viewportState.invert === 'boolean' && typeof viewport?.setProperties === 'function') {
      const volumeId = getViewportVolumeId(viewport, viewportState.displaySetInstanceUID);
      if (volumeId) {
        viewport.setProperties({ invert: viewportState.invert }, volumeId);
      } else {
        viewport.setProperties({ invert: viewportState.invert });
      }
    }

    if (Object.keys(presentationPatch).length || typeof viewportState.invert === 'boolean') {
      viewport?.render?.();
    }
  };

  const applyCameraState = (viewport : ViewPort, viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]) => {
    if (typeof viewport?.setCamera !== 'function') {
      return;
    }

    const camera = viewportState.camera;
    if (camera && typeof camera === 'object') {
      viewport.setCamera(camera);
      viewport.render?.();
      return;
    }

    const cameraPatch: Record<string, unknown> = {};
    if (typeof viewportState.zoom === 'number' && viewportState.zoom !== 0) {
      cameraPatch.parallelScale = 1 / viewportState.zoom;
    }
    if (viewportState.pan && typeof viewportState.pan.x === 'number' && typeof viewportState.pan.y === 'number') {
      cameraPatch.focalPoint = [viewportState.pan.x, viewportState.pan.y, 0];
    }
    if (Object.keys(cameraPatch).length) {
      viewport.setCamera(cameraPatch);
      viewport.render?.();
    }
  };

  const getGridViewport = (viewportId: string) => {
    const viewports = services().viewportGridService?.getState?.()?.viewports;
    return viewports?.get?.(viewportId) ?? null;
  };

  const getDisplaySet = (viewportId: string) => {
    const displaySetInstanceUID = getGridViewport(viewportId)?.displaySetInstanceUIDs?.[0];
    if (!displaySetInstanceUID) {
      return null;
    }

    try {
      return services().displaySetService?.getDisplaySetByUID?.(displaySetInstanceUID) ?? null;
    } catch {
      return null;
    }
  };

  const getFirstViewportId = () => {
    const viewports = services().viewportGridService?.getState?.()?.viewports;
    return Array.from(viewports?.keys?.() ?? [])[0] as string | undefined;
  };

  const resolveApplyViewportId = (requestedViewportId?: string) => {
    const { cornerstoneViewportService } = services();
    const candidates = [requestedViewportId, getActiveViewportId(), 'default', getFirstViewportId()].filter(Boolean);

    return candidates.find(viewportId =>
      cornerstoneViewportService?.getCornerstoneViewport?.(viewportId as string)
    ) as string | undefined;
  };

  const resolveTargetDisplaySet = (viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number]) => {
    const { displaySetService } = services();

    if (viewportState.displaySetInstanceUID) {
      const displaySet = displaySetService?.getDisplaySetByUID?.(viewportState.displaySetInstanceUID);
      if (displaySet) {
        return displaySet;
      }
    }

    const seriesInstanceUID = viewportState.seriesInstanceUID;
    if (!seriesInstanceUID) {
      return null;
    }

    const displaySets = displaySetService?.getDisplaySetsForSeries?.(seriesInstanceUID) ?? [];
    return displaySets.find((displaySet: any) => !displaySet.isOverlayDisplaySet) ?? displaySets[0] ?? null;
  };

  const waitForAppliedImageIndex = (viewportId: string, imageIndex: number, sequence?: number, timeoutMs = 1500) =>
    new Promise<'applied' | 'stale' | 'timeout'>(resolve => {
      const startedAt = Date.now();
      const check = () => {
        if (isStaleApplySequence(sequence)) {
          resolve('stale');
          return;
        }

        const viewport = services().cornerstoneViewportService?.getCornerstoneViewport?.(viewportId);
        if (viewport && getImageIndex(viewport) === imageIndex) {
          resolve('applied');
          return;
        }

        if (Date.now() - startedAt >= timeoutMs) {
          resolve('timeout');
          return;
        }

        window.setTimeout(check, 50);
      };

      check();
    });

  const waitForAppliedWindowLevel = (
    viewportId: string,
    viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number],
    timeoutMs = 500
  ) =>
    new Promise<boolean>(resolve => {
      const targetVOIRange = getTargetVOIRange(viewportState);
      if (!targetVOIRange) {
        resolve(true);
        return;
      }

      const startedAt = Date.now();
      const check = () => {
        const viewport = services().cornerstoneViewportService?.getCornerstoneViewport?.(viewportId);
        const { voiRange } = getWindowLevelState(viewport, viewportState.displaySetInstanceUID);
        if (areVOIRangesClose(voiRange, targetVOIRange)) {
          resolve(true);
          return;
        }

        if (Date.now() - startedAt >= timeoutMs) {
          resolve(false);
          return;
        }

        window.setTimeout(check, 50);
      };

      check();
    });

  const waitForViewportDisplaySet = (viewportId: string, displaySetInstanceUID: string, timeoutMs = 1500) =>
    new Promise<boolean>(resolve => {
      const startedAt = Date.now();
      const check = () => {
        const displaySetInstanceUIDs = getGridViewport(viewportId)?.displaySetInstanceUIDs ?? [];
        if (displaySetInstanceUIDs[0] === displaySetInstanceUID) {
          resolve(true);
          return;
        }

        if (Date.now() - startedAt >= timeoutMs) {
          resolve(false);
          return;
        }

        window.setTimeout(check, 50);
      };

      check();
    });

  const getInstance = (displaySet: any, currentImageId: string | null, imageIndex: number) => {
    const instances = displaySet?.instances ?? displaySet?.images ?? [];
    if (!instances.length) {
      return null;
    }

    return instances.find((instance: any) => instance?.imageId === currentImageId) ?? instances[imageIndex] ?? null;
  };

  const captureViewportState = (viewportId: string) => {
    const { cornerstoneViewportService } = services();
    const viewport = cornerstoneViewportService.getCornerstoneViewport(viewportId);
    if (!viewport) {
      return null;
    }

    const imageIndex = getImageIndex(viewport);
    const imageIds = getImageIds(viewport);
    const currentImageId = getCurrentImageId(viewport, imageIds, imageIndex);
    const displaySet = getDisplaySet(viewportId);
    if (!displaySet) {
      return null;
    }

    const instance = getInstance(displaySet, currentImageId, imageIndex);
    const numberOfSlices = getNumberOfSlices(viewport, imageIds, displaySet);
    const studyInstanceUID = instance?.StudyInstanceUID ?? displaySet.StudyInstanceUID ?? '';
    const seriesInstanceUID = instance?.SeriesInstanceUID ?? displaySet.SeriesInstanceUID ?? '';
    const sopInstanceUID = instance?.SOPInstanceUID ?? displaySet.SOPInstanceUID ?? null;
    const { camera, zoom, pan } = getZoomPanState(viewport);
    const { voiRange, windowLevel } = getWindowLevelState(viewport, displaySet.displaySetInstanceUID);
    const { rotation, flipHorizontal, flipVertical, invert } = getPresentationState(viewport, displaySet.displaySetInstanceUID);
    const { isPlaying, frameRate } = getCineState(viewportId);

    if (!imageIds.length && !displaySet?.instances?.length && !displaySet?.images?.length) {
      return null;
    }
    if (!studyInstanceUID || !seriesInstanceUID) {
      return null;
    }

    return {
      viewportId,
      displaySetInstanceUID: displaySet.displaySetInstanceUID,
      studyInstanceUID,
      seriesInstanceUID,
      sopInstanceUID,
      imageIndex,
      numberOfSlices,
      zoom,
      pan,
      camera,
      voiRange,
      windowLevel,
      rotation,
      flipHorizontal,
      flipVertical,
      invert,
      isPlaying,
      frameRate,
    };
  };

  const getCurrentLayoutSignature = () => {
    const { viewportGridService } = services();
    const gridState = viewportGridService?.getState?.();
    const layoutState = getLayoutState();
    return JSON.stringify({
      rows: layoutState.layout?.numRows ?? null,
      cols: layoutState.layout?.numCols ?? null,
      type: layoutState.layout?.layoutType ?? null,
      hp: layoutState.isHangingProtocolLayout ? layoutState.hangingProtocol?.protocolId ?? null : null,
      stage: layoutState.isHangingProtocolLayout ? layoutState.hangingProtocol?.stageIndex ?? null : null,
      count: gridState?.viewports?.size ?? 0,
    });
  };

  const captureActiveSliceState = (forceFullLayout = false): CaptureResult => {
    lastCaptureAttemptAt = Date.now();

    const { viewportGridService, cornerstoneViewportService, displaySetService } = services();
    if (!runtime.servicesManager?.services) {
      return { reason: 'services_unavailable' };
    }
    if (!viewportGridService?.getState || !viewportGridService?.getActiveViewportId) {
      return { reason: 'viewport_grid_unavailable' };
    }
    if (!cornerstoneViewportService?.getCornerstoneViewport) {
      return { reason: 'cornerstone_viewport_service_unavailable' };
    }
    if (!displaySetService?.getDisplaySetByUID) {
      return { reason: 'display_set_service_unavailable' };
    }

    const viewportId = getActiveViewportId();
    if (!viewportId) {
      return { reason: 'active_viewport_missing' };
    }

    const activeViewportState = captureViewportState(viewportId);
    if (!activeViewportState) {
      return { reason: 'cornerstone_viewport_missing' };
    }

    const gridState = viewportGridService.getState();
    const currentLayoutSignature = getCurrentLayoutSignature();
    const layoutChanged = currentLayoutSignature !== lastCapturedLayoutSignature;
    const shouldCaptureAllViewports = forceFullLayout || layoutChanged;
    const viewportIds = shouldCaptureAllViewports
      ? Array.from(gridState?.viewports?.keys?.() ?? []) as string[]
      : [viewportId];
    const viewports = viewportIds
      .map(captureViewportState)
      .filter((state): state is NonNullable<ReturnType<typeof captureViewportState>> => !!state);

    lastCapturedLayoutSignature = currentLayoutSignature;

    return { state: {
      studyInstanceUID: activeViewportState.studyInstanceUID,
      seriesInstanceUID: activeViewportState.seriesInstanceUID,
      sopInstanceUID: activeViewportState.sopInstanceUID,
      viewportLayout: getViewportLayout(),
      ...getLayoutState(),
      layoutViewportCount: gridState?.viewports?.size ?? viewports.length,
      activeViewportId: viewportId,
      viewports: viewports.length ? viewports : [activeViewportState],
    } };
  };

  const emitViewerStateChanged = (forceFullLayout = false) => {
    if (emitViewerStateTimer !== null) {
      window.clearTimeout(emitViewerStateTimer);
      emitViewerStateTimer = null;
    }

    if (mode !== 'presenter') {
      return;
    }

    const result = captureActiveSliceState(forceFullLayout);
    if (!result.state) {
      emitCaptureFailure(result.reason);
      return;
    }

    lastCaptureFailure = '';
    lastCaptureResult = 'success';
    lastCaptureError = null;
    const { state } = result;
    const serialized = JSON.stringify(state);
    if (serialized === lastViewerState) {
      return;
    }

    lastViewerState = serialized;
    post({ type: 'viewer_state_changed', state });
  };

  const scheduleViewerStateChanged = () => {
    if (mode !== 'presenter' || emitViewerStateTimer !== null) {
      return;
    }

    emitViewerStateTimer = window.setTimeout(() => {
      emitViewerStateTimer = null;
      emitViewerStateChanged(false);
    }, 75);
  };

  const failApply = (sequence: number | undefined, reason: string) => {
    post({ type: 'remote_state_apply_failed', sequence, reason });
  };

  const emitPresenterPointer = (event: MouseEvent, viewportId: string) => {
    if (mode !== 'presenter') {
      return;
    }

    const now = Date.now();
    if (now - lastPointerSentAt < 50) {
      return;
    }

    const element = event.currentTarget as HTMLElement | null;
    const rect = element?.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      return;
    }

    const x = (event.clientX - rect.left) / rect.width;
    const y = (event.clientY - rect.top) / rect.height;
    if (x < 0 || x > 1 || y < 0 || y > 1) {
      return;
    }

    lastPointerSentAt = now;
    post({ type: 'viewer_pointer_changed', viewportId, x, y });
  };

  const isStaleApplySequence = (sequence: number | undefined) =>
    typeof sequence === 'number' && sequence < latestApplySequence;

  const waitForViewportCount = (count: number, timeoutMs = 1500) =>
    new Promise<boolean>(resolve => {
      const startedAt = Date.now();
      const check = () => {
        const viewportCount = services().viewportGridService?.getState?.()?.viewports?.size ?? 0;
        if (viewportCount >= count) {
          resolve(true);
          return;
        }

        if (Date.now() - startedAt >= timeoutMs) {
          resolve(false);
          return;
        }

        window.setTimeout(check, 50);
      };

      check();
    });

  const getLayoutSignature = (state: NonNullable<ParentMessage['state']>) => {
    const layout = (state as any).layout;
    const hangingProtocol = (state as any).hangingProtocol;
    return JSON.stringify({
      rows: layout?.numRows ?? null,
      cols: layout?.numCols ?? null,
      type: layout?.layoutType ?? null,
      hp: (state as any).isHangingProtocolLayout ? hangingProtocol?.protocolId ?? null : null,
      stage: (state as any).isHangingProtocolLayout ? hangingProtocol?.stageIndex ?? null : null,
      count: (state as any).layoutViewportCount ?? state.viewports?.length ?? 0,
    });
  };

  const applyLayoutState = async (state: NonNullable<ParentMessage['state']>, sequence?: number) => {
    const { viewportGridService } = services();
    const targetLayout = (state as any).layout;
    const nextLayoutSignature = getLayoutSignature(state);
    if (nextLayoutSignature === lastAppliedLayoutSignature) {
      return { ready: true, changed: false };
    }

    const targetViewportCount = (state as any).layoutViewportCount ?? state.viewports?.length ?? 1;
    if (!viewportGridService?.getState || !targetLayout) {
      lastAppliedLayoutSignature = nextLayoutSignature;
      return { ready: true, changed: false };
    }

    const currentLayout = viewportGridService.getState()?.layout;
    const currentViewportCount = viewportGridService.getState()?.viewports?.size ?? 0;
    const alreadyMatches =
      currentLayout?.numRows === targetLayout.numRows &&
      currentLayout?.numCols === targetLayout.numCols &&
      (currentLayout?.layoutType || 'grid') === (targetLayout.layoutType || 'grid') &&
      currentViewportCount >= targetViewportCount;

    if (alreadyMatches) {
      lastAppliedLayoutSignature = nextLayoutSignature;
      return { ready: true, changed: false };
    }

    const hangingProtocol = (state as any).hangingProtocol;
    const canRunCommand = typeof runtime.commandsManager?.run === 'function';
    if ((state as any).isHangingProtocolLayout && hangingProtocol?.protocolId && hangingProtocol.protocolId !== 'default' && canRunCommand) {
      runtime.commandsManager!.run({
        commandName: 'setHangingProtocol',
        commandOptions: {
          protocolId: hangingProtocol.protocolId,
          stageIndex: hangingProtocol.stageIndex,
          StudyInstanceUID: state.studyInstanceUID,
          reset: true,
        },
      });
    } else if (canRunCommand) {
      runtime.commandsManager!.run({
        commandName: 'setViewportGridLayout',
        commandOptions: {
          numRows: targetLayout.numRows,
          numCols: targetLayout.numCols,
        },
      });
    }

    const ready = await waitForViewportCount(targetViewportCount);
    if (isStaleApplySequence(sequence)) {
      return { ready: false, changed: true };
    }

    updateReadOnlyGuard();
    if (ready) {
      lastAppliedLayoutSignature = nextLayoutSignature;
    }
    return { ready, changed: true };
  };

  const applyViewportVisualState = async (
    viewportState: NonNullable<NonNullable<ParentMessage['state']>['viewports']>[number],
    sequence: number | undefined
  ) => {
    if (!viewportState) {
      return 'viewport_not_ready';
    }

    if (typeof viewportState.imageIndex !== 'number') {
      return 'image_index_missing';
    }

    const { viewportGridService, cornerstoneViewportService } = services();
    if (!viewportGridService?.getState || !cornerstoneViewportService?.getCornerstoneViewport) {
      return 'viewport_not_ready';
    }

    const viewportId = resolveApplyViewportId(viewportState.viewportId);
    if (!viewportId) {
      return 'viewport_not_ready';
    }

    const targetDisplaySet = resolveTargetDisplaySet(viewportState);
    const currentDisplaySet = getDisplaySet(viewportId);
    if (!targetDisplaySet) {
      return viewportState.seriesInstanceUID ? 'series_not_found' : 'display_set_missing';
    }

    const imageIndex = viewportState.imageIndex;
    const numberOfSlices = viewportState.numberOfSlices || targetDisplaySet.numImageFrames || targetDisplaySet.instances?.length;
    if (typeof numberOfSlices === 'number' && (imageIndex < 0 || imageIndex >= numberOfSlices)) {
      return 'image_index_out_of_range';
    }

    try {
      if (currentDisplaySet?.displaySetInstanceUID !== targetDisplaySet.displaySetInstanceUID) {
        viewportGridService.setDisplaySetsForViewport({
          viewportId,
          displaySetInstanceUIDs: [targetDisplaySet.displaySetInstanceUID],
        });

        const switched = await waitForViewportDisplaySet(viewportId, targetDisplaySet.displaySetInstanceUID);
        if (isStaleApplySequence(sequence)) {
          return 'stale';
        }
        if (!switched) {
          return 'display_set_switch_failed';
        }
      }

      const viewport = cornerstoneViewportService.getCornerstoneViewport(viewportId);
      if (!viewport?.element) {
        return 'viewport_not_ready_after_series_switch';
      }

      const currentImageIndex = getImageIndex(viewport);
      if (currentImageIndex === imageIndex) {
        applyCameraState(viewport, viewportState);
        applyWindowLevelState(viewport, viewportState);
        applyPresentationState(viewport, viewportState);
        applyCineState(viewportId, viewport.element, viewportState);
        const windowLevelApplied = await waitForAppliedWindowLevel(viewportId, viewportState);
        if (isStaleApplySequence(sequence)) {
          return 'stale';
        }
        if (!windowLevelApplied) {
          return 'window_level_apply_failed';
        }
        return null;
      }

      if (isStaleApplySequence(sequence)) {
        return 'stale';
      }
      utilities.jumpToSlice(viewport.element, { imageIndex });
      const applied = await waitForAppliedImageIndex(viewportId, imageIndex, sequence);
      if (applied === 'stale') {
        return 'stale';
      }
      if (applied === 'timeout') {
        return 'apply_failed';
      }

      const updatedViewport = cornerstoneViewportService.getCornerstoneViewport(viewportId);
      applyCameraState(updatedViewport, viewportState);
      applyWindowLevelState(updatedViewport, viewportState);
      applyPresentationState(updatedViewport, viewportState);
      if (updatedViewport?.element) {
        applyCineState(viewportId, updatedViewport.element, viewportState);
      }
      const windowLevelApplied = await waitForAppliedWindowLevel(viewportId, viewportState);
      if (isStaleApplySequence(sequence)) {
        return 'stale';
      }
      if (!windowLevelApplied) {
        return 'window_level_apply_failed';
      }

      return null;
    } catch {
      return 'apply_failed';
    }
  };

  const applyRemoteViewerState = async (message: ParentMessage) => {
    const sequence = message.sequence;
    if (typeof sequence === 'number') {
      if (sequence < latestApplySequence) {
        return;
      }
      latestApplySequence = sequence;
    }

    if (!message.state?.viewports?.length) {
      failApply(sequence, 'viewport_not_ready');
      return;
    }

    const layoutApplyResult = await applyLayoutState(message.state, sequence);
    if (!layoutApplyResult.ready) {
      failApply(sequence, 'layout_apply_failed');
      return;
    }

    const activeViewportState = message.state.viewports.find(viewport => viewport.viewportId === message.state?.activeViewportId) ?? message.state.viewports[0];
    const orderedViewports = layoutApplyResult.changed
      ? [
          ...message.state.viewports.filter(viewport => viewport !== activeViewportState),
          activeViewportState,
        ]
      : [activeViewportState];

    for (const viewportState of orderedViewports) {
      const failure = await applyViewportVisualState(viewportState, sequence);
      if (failure === 'stale') {
        return;
      }
      if (failure) {
        failApply(sequence, failure);
        return;
      }
    }

    post({ type: 'remote_state_applied', sequence });
  };

  const enqueueRemoteViewerStateApply = (message: ParentMessage) => {
    const sequence = message.sequence;
    if (typeof sequence === 'number') {
      if (sequence < latestApplySequence) {
        return;
      }
      latestApplySequence = sequence;
    }

    pendingRemoteViewerState = message;
    if (remoteViewerStateApplyInFlight) {
      return;
    }

    remoteViewerStateApplyInFlight = true;
    void (async () => {
      try {
        while (pendingRemoteViewerState) {
          const nextMessage = pendingRemoteViewerState;
          pendingRemoteViewerState = null;
          await applyRemoteViewerState(nextMessage);
        }
      } finally {
        remoteViewerStateApplyInFlight = false;
        if (pendingRemoteViewerState) {
          enqueueRemoteViewerStateApply(pendingRemoteViewerState);
        }
      }
    })();
  };

  const onParentMessage = (event: MessageEvent<ParentMessage>) => {
    const message = event.data;
    if (!message || typeof message !== 'object' || !message.type) {
      return;
    }
    if (event.source !== window.parent || event.origin !== window.location.origin) {
      return;
    }

    switch (message.type) {
      case 'set_presenter_mode':
        mode = 'presenter';
        seedKnownMeasurements();
        updateReadOnlyGuard();
        removePresenterPointerOverlay();
        emitViewerStateChanged();
        window.setTimeout(emitViewerStateChanged, 500);
        window.setTimeout(emitViewerStateChanged, 1500);
        window.setTimeout(emitViewerStateChanged, 3000);
        break;
      case 'set_synced_mode':
      case 'set_read_only':
        mode = 'synced';
        updateReadOnlyGuard();
        break;
      case 'start_local_inspect':
        mode = 'local_inspect';
        updateReadOnlyGuard();
        removePresenterPointerOverlay();
        break;
      case 'return_to_synced_view':
        mode = 'synced';
        updateReadOnlyGuard();
        enqueueRemoteViewerStateApply(message);
        break;
      case 'apply_remote_viewer_state':
        enqueueRemoteViewerStateApply(message);
        break;
      case 'show_presenter_pointer':
        applyPresenterPointer(message.pointer);
        break;
      case 'apply_measurement_artifact':
        applyMeasurementArtifact(message);
        break;
      case 'clear_session':
        mode = 'disconnected';
        updateReadOnlyGuard();
        removePresenterPointerOverlay();
        break;
    }
  };

  const bindViewportEvents = () => {
    const { cornerstoneViewportService, viewportGridService } = services();
    const state = viewportGridService?.getState?.();
    const viewportIds = Array.from(state?.viewports?.keys?.() ?? []) as string[];
    const eventNames = [
      EVENTS.STACK_NEW_IMAGE,
      EVENTS.STACK_VIEWPORT_SCROLL,
      EVENTS.VOLUME_NEW_IMAGE,
      EVENTS.VIEWPORT_NEW_IMAGE_SET,
      EVENTS.CAMERA_MODIFIED,
      Enums.Events.VOI_MODIFIED,
    ].filter(Boolean);
    const cineService = services().cineService;

    const viewportsWithElements = viewportIds.map(viewportId => ({
      viewportId,
      element: cornerstoneViewportService?.getCornerstoneViewport?.(viewportId)?.element,
    }));
    const key = viewportsWithElements
      .map(({ viewportId, element }) => `${viewportId}:${element ? 'ready' : 'pending'}`)
      .join('|');

    if (key === boundViewportIds) {
      return;
    }

    viewportEventDisposers.forEach(dispose => dispose());
    viewportEventDisposers = [];
    boundViewportIds = key;

    viewportsWithElements.forEach(({ viewportId, element }) => {
      if (!element) {
        return;
      }

      const pointerListener = (event: MouseEvent) => emitPresenterPointer(event, viewportId);
      element.addEventListener('mousemove', pointerListener);
      viewportEventDisposers.push(() => element.removeEventListener('mousemove', pointerListener));

      eventNames.forEach(eventName => {
        element.addEventListener(eventName, scheduleViewerStateChanged);
        viewportEventDisposers.push(() => element.removeEventListener(eventName, scheduleViewerStateChanged));
      });
    });

    if (viewportGridService?.EVENTS && typeof viewportGridService.subscribe === 'function') {
      [viewportGridService.EVENTS.LAYOUT_CHANGED, viewportGridService.EVENTS.GRID_STATE_CHANGED]
        .filter(Boolean)
        .forEach(eventName => {
          const subscription = viewportGridService.subscribe(eventName, () => emitViewerStateChanged(true));
          viewportEventDisposers.push(() => subscription.unsubscribe());
        });
    }

    if (cineService?.EVENTS?.CINE_STATE_CHANGED && typeof cineService.subscribe === 'function') {
      const subscription = cineService.subscribe(cineService.EVENTS.CINE_STATE_CHANGED, emitViewerStateChanged);
      viewportEventDisposers.push(() => subscription.unsubscribe());
    }

    const measurementService = services().measurementService;
    if (measurementService?.EVENTS && typeof measurementService.subscribe === 'function') {
      [measurementService.EVENTS.MEASUREMENT_ADDED, measurementService.EVENTS.RAW_MEASUREMENT_ADDED]
        .filter(Boolean)
        .forEach(eventName => {
          const subscription = measurementService.subscribe(eventName, emitMeasurementCreated);
          viewportEventDisposers.push(() => subscription.unsubscribe());
        });
      if (measurementService.EVENTS.MEASUREMENT_UPDATED) {
        const subscription = measurementService.subscribe(measurementService.EVENTS.MEASUREMENT_UPDATED, emitMeasurementUpdated);
        viewportEventDisposers.push(() => subscription.unsubscribe());
      }
      if (measurementService.EVENTS.MEASUREMENT_REMOVED) {
        const subscription = measurementService.subscribe(measurementService.EVENTS.MEASUREMENT_REMOVED, emitMeasurementRemoved);
        viewportEventDisposers.push(() => subscription.unsubscribe());
      }
    }

    const toolsEvents = ToolsEnums?.Events;
    if (toolsEvents) {
      const rawAnnotationListeners: Array<[string | undefined, EventListener]> = [
        [toolsEvents.ANNOTATION_COMPLETED, emitMeasurementCreated as EventListener],
        [toolsEvents.ANNOTATION_MODIFIED, emitMeasurementUpdated as EventListener],
        [toolsEvents.ANNOTATION_REMOVED, emitMeasurementRemoved as EventListener],
      ];

      rawAnnotationListeners
        .filter((entry): entry is [string, EventListener] => typeof entry[0] === 'string')
        .forEach(([eventName, listener]) => {
          eventTarget.addEventListener(eventName, listener);
          viewportEventDisposers.push(() => eventTarget.removeEventListener(eventName, listener));
        });
    }
  };

  window.addEventListener('message', onParentMessage);
  document.addEventListener('keydown', blockSyncedUserInput, true);
  document.addEventListener('keyup', blockSyncedUserInput, true);
  post({ type: 'viewer_ready' });
  post({ type: 'bridge_capabilities', capabilities: getCapabilities() });

  const statusInterval = window.setInterval(() => {
    post({
      type: 'bridge_status',
      status: {
        source: 'ohif-extension',
        mode,
        capabilities: getCapabilities(),
        capture: getCaptureDiagnostics(),
      },
    });
  }, 1000);

  const bindInterval = window.setInterval(bindViewportEvents, 1000);
  const readOnlyOverlayInterval = window.setInterval(updateReadOnlyGuard, 500);
  const presenterCaptureInterval = window.setInterval(emitViewerStateChanged, 500);

  bindViewportEvents();

  return () => {
    window.removeEventListener('message', onParentMessage);
    document.removeEventListener('keydown', blockSyncedUserInput, true);
    document.removeEventListener('keyup', blockSyncedUserInput, true);
    window.clearInterval(statusInterval);
    window.clearInterval(bindInterval);
    window.clearInterval(readOnlyOverlayInterval);
    window.clearInterval(presenterCaptureInterval);
    if (emitViewerStateTimer !== null) {
      window.clearTimeout(emitViewerStateTimer);
      emitViewerStateTimer = null;
    }
    removeReadOnlyOverlays();
    removePresenterPointerOverlay();
    viewportEventDisposers.forEach(dispose => dispose());
    viewportEventDisposers = [];
  };
};

let disposeBridge: undefined | (() => void);

const extension = {
  id,
  onModeEnter({ servicesManager, commandsManager }: withAppTypes) {
    if (!isBridgeEnabled() || disposeBridge) {
      return;
    }

    disposeBridge = createBridge({
      servicesManager,
      commandsManager,
    });
  },
  onModeExit() {
    disposeBridge?.();
    disposeBridge = undefined;
  },
};

export default extension;
