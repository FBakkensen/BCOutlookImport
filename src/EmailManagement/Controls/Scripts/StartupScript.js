// Ensure proper initialization with multiple approaches
(function() {
    // Define our initialization function
    function initializeDropZone() {
        console.log("Initializing Outlook Email Drop control");

        // Clear any existing content first
        document.body.innerHTML = '';

        // Force body to take full size of the container
        document.body.style.margin = '0';
        document.body.style.padding = '0';
        document.body.style.width = '100%';
        document.body.style.height = '100%';
        document.body.style.overflow = 'hidden';
        document.body.style.display = 'block';

        // Create a fallback message if InitializeControl isn't available
        if (typeof InitializeControl === 'function') {
            InitializeControl();
        } else {
            console.error("InitializeControl function not found - creating fallback element");
            createFallbackControl();
        }

        // Setup visibility check interval
        setInterval(function() {
            const dropZone = document.getElementById('dropZone');
            if (!dropZone || dropZone.offsetHeight < 100) {
                console.log("Reinitializing drop zone due to visibility issues");

                // Clear and reinitialize if not visible
                if (typeof InitializeControl === 'function') {
                    document.body.innerHTML = '';
                    InitializeControl();
                } else {
                    createFallbackControl();
                }
            }

            // Force height reporting
            if (typeof GetControlHeight === 'function') {
                GetControlHeight();
            } else if (typeof Microsoft !== 'undefined' &&
                       typeof Microsoft.Dynamics !== 'undefined' &&
                       typeof Microsoft.Dynamics.NAV !== 'undefined') {
                // Try to report height directly
                Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlHeightReturned', [200]);
            }
        }, 1000);
    }

    // Fallback control when initialization fails
    function createFallbackControl() {
        const fallback = document.createElement('div');
        fallback.id = 'dropZone'; // Use same ID for consistency
        fallback.style.width = '100%';
        fallback.style.height = '100px';
        fallback.style.border = '3px dashed #0078d4';
        fallback.style.borderRadius = '5px';
        fallback.style.display = 'flex';
        fallback.style.alignItems = 'center';
        fallback.style.justifyContent = 'center';
        fallback.style.flexDirection = 'column';
        fallback.style.backgroundColor = '#e6f2ff';
        fallback.style.margin = '5px 0';
        fallback.style.padding = '10px';
        fallback.style.boxSizing = 'border-box';
        fallback.style.textAlign = 'center';
        fallback.style.fontSize = '14px';
        fallback.style.fontWeight = 'bold';

        fallback.innerHTML = `
            <div style="font-size: 24px; color: #0078d4; margin-bottom: 5px;">📧</div>
            <div style="color: #0e0e0e; font-weight: 600;">Drop email here</div>
        `;

        document.body.appendChild(fallback);

        // Try to notify AL that control is ready if Microsoft.Dynamics.NAV is available
        if (typeof Microsoft !== 'undefined' &&
            typeof Microsoft.Dynamics !== 'undefined' &&
            typeof Microsoft.Dynamics.NAV !== 'undefined') {
            Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlReady', []);

            // Also report a smaller height
            Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlHeightReturned', [120]);
        }
    }

    // Try multiple approaches to ensure initialization

    // Method 1: Standard DOMContentLoaded
    document.addEventListener("DOMContentLoaded", function() {
        setTimeout(initializeDropZone, 100);
    });

    // Method 2: Use load event as backup
    window.addEventListener("load", function() {
        setTimeout(initializeDropZone, 200);
    });

    // Method 3: Forceful backup in case the events don't fire
    setTimeout(function() {
        if (!document.getElementById('dropZone')) {
            console.log("Forcing initialization after timeout");
            initializeDropZone();
        }
    }, 500);
})();