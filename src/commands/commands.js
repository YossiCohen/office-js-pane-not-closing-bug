/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}


function ribbonItemEnable(ribbonButtonId, state){
  Office.ribbon.requestUpdate({
    tabs: [
        {
            id: "MyTab1", 
            groups: [
                {
                  id: "FirstGroup",
                  controls: [
                    {
                        id: ribbonButtonId, 
                        enabled: state
                    }
                  ]
                }
            ]
        }
    ]
});
}

const g = getGlobal();

