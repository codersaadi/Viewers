/** @type {AppTypes.Config} */
window.config = {
  routerBasename: '/ohif',
  whiteLabeling: {
    createLogoComponentFn: function (React) {
      return React.createElement(
        'div',
        {
          className: 'flex items-center text-base font-semibold tracking-[0.18em] text-primary-light',
        },
        'LUMEX'
      );
    },
  },
  extensions: ['@lumex/extension-meeting'],
  modes: [],
  showStudyList: false, // Lumex manages its own study list
  maxNumberOfWebWorkers: navigator.hardwareConcurrency || 4,
  showLoadingIndicator: true,
  showWarningMessageForCrossOrigin: false,
  showCPUFallbackMessage: false,
  strictZSpacingForVolumeViewport: true,
  defaultDataSourceName: 'orthanc',
  dataSources: [
    {
      namespace: '@ohif/extension-default.dataSourcesModule.dicomweb',
      sourceName: 'orthanc',
      configuration: {
        friendlyName: 'Lumex Orthanc',
        name: 'orthanc',
        // Use Next.js proxy which handles authentication
        wadoUriRoot: '/api/orthanc/dicom-web',
        qidoRoot: '/api/orthanc/dicom-web',
        wadoRoot: '/api/orthanc/dicom-web',
        qidoSupportsIncludeField: true,
        supportsReject: false,
        dicomUploadEnabled: false, // We handle uploads separately
        imageRendering: 'wadors',
        thumbnailRendering: 'wadors',
        enableStudyLazyLoad: true,
        supportsFuzzyMatching: true,
        supportsWildcard: true,
        omitQuotationForMultipartRequest: true,
        bulkDataURI: {
          enabled: true,
          // Orthanc returns BulkDataURIs pointing to http://localhost:8042/dicom-web/...
          // Rewrite them to go through our Next.js proxy instead
          startsWith: '',
          prefixWith: '/api/orthanc/dicom-web/',
        },
      },
    },
  ],
  // User preferences
  userPreferences: {
    // Show patient info in sidebar
    showPatientInfo: 'hidden',
  },
  // Performance optimizations
  studyPrefetcher: {
    enabled: true,
    displaySetsCount: 2,
    maxNumPrefetchRequests: 10,
    order: 'closest',
  },
};
