// Points the page at the in-browser backend (Code.gs + mock-gas.js). Test use only.
(() => {
  'use strict';
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CREW_CODE')) setUp();
  window.__mock.crew = props.getProperty('CREW_CODE');
  window.__mock.host = props.getProperty('HOST_CODE');
  window.__mock.calls = [];
  // Arrive the way a friend would: by the link with the crew code on the end.
  if (!location.hash && !/nocode/.test(location.search)) location.hash = window.__mock.crew;

  window.__secondSeatTransport = payload => new Promise((resolve, reject) => {
    window.__mock.calls.push(payload);
    setTimeout(() => {
      if (window.__mock.offline) return reject(new Error('offline'));
      const out = doPost({ postData: { contents: JSON.stringify(payload) } });
      resolve(JSON.parse(out.getContent()));
    }, 120);
  });
})();
