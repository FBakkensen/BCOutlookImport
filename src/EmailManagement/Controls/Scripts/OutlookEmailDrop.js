var customerId = '';

function InitializeControl() {
    // Create a container div with more compact dimensions
    const containerDiv = document.createElement('div');
    containerDiv.style.width = '100%';
    containerDiv.style.height = '100%';
    containerDiv.style.minHeight = '100px'; // Reduced minimum height
    containerDiv.style.position = 'absolute';
    containerDiv.style.top = '0';
    containerDiv.style.left = '0';
    containerDiv.style.padding = '3px';
    containerDiv.style.boxSizing = 'border-box';
    containerDiv.style.backgroundColor = '#ffffff';
    containerDiv.style.zIndex = '1000';

    // Create the drop zone
    const dropZone = document.createElement('div');
    dropZone.id = 'dropZone';
    dropZone.className = 'outlook-drop-zone';

    // Create a simpler SVG icon that's smaller
    const iconSvg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="17 8 12 3 7 8"></polyline>
        <line x1="12" y1="3" x2="12" y2="15"></line>
    </svg>`;

    // Create more compact HTML content
    dropZone.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; width: 100%;">
            <div class="drop-icon">${iconSvg}</div>
            <div class="drop-message">Drop email here</div>
        </div>
    `;

    // Set a subtle pulsing animation
    const style = document.createElement('style');
    style.textContent = `
        @keyframes pulse-border {
            0% { border-color: #0078d4; }
            50% { border-color: #69b0ff; }
            100% { border-color: #0078d4; }
        }
        #dropZone {
            animation: pulse-border 2s infinite;
        }
    `;
    document.head.appendChild(style);

    // Setup event listeners
    dropZone.addEventListener('dragover', handleDragOver);
    dropZone.addEventListener('dragleave', handleDragLeave);
    dropZone.addEventListener('drop', handleFileDrop);
    dropZone.addEventListener('click', handleClick);

    // Append elements
    containerDiv.appendChild(dropZone);
    document.body.appendChild(containerDiv);

    // Notify AL that the control is ready
    Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlReady', []);

    // Force height reporting after a short delay
    setTimeout(GetControlHeight, 200);
}

function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    this.classList.add('drag-over');

    // Update the drop message to provide clear instructions
    const dropMessage = this.querySelector('.drop-message');
    if (dropMessage) {
        dropMessage.textContent = 'Release to import email';
        dropMessage.style.fontWeight = 'bold';
    }

    // Make the icon larger and animate it slightly
    const dropIcon = this.querySelector('.drop-icon');
    if (dropIcon) {
        dropIcon.style.fontSize = '30px';
        dropIcon.style.transition = 'all 0.2s ease';
    }
}

function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    this.classList.remove('drag-over');

    // Reset the drop message
    const dropMessage = this.querySelector('.drop-message');
    if (dropMessage) {
        dropMessage.textContent = 'Drop Outlook email here';
        dropMessage.style.fontWeight = '500';
    }

    // Reset the icon
    const dropIcon = this.querySelector('.drop-icon');
    if (dropIcon) {
        dropIcon.style.fontSize = '24px';
    }
}

function handleClick(event) {
    event.preventDefault();
    // Create a file input element
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.eml,.msg';
    fileInput.style.display = 'none';

    fileInput.onchange = function(e) {
        const file = e.target.files[0];
        if (file) {
            processFile(file);
        }
        document.body.removeChild(fileInput);
    };

    document.body.appendChild(fileInput);
    fileInput.click();
}

function handleFileDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    this.classList.remove('drag-over');

    // Reset the drop message
    const dropMessage = this.querySelector('.drop-message');
    if (dropMessage) {
        dropMessage.textContent = 'Drop Outlook email here';
        dropMessage.style.fontWeight = '500';
    }

    // Reset the icon
    const dropIcon = this.querySelector('.drop-icon');
    if (dropIcon) {
        dropIcon.style.fontSize = '24px';
    }

    // Check if files were dropped
    if (event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files.length > 0) {
        const file = event.dataTransfer.files[0];
        processFile(file);
    }
}

function processFile(file) {
    // Check if file is an email (eml or msg)
    const validExtensions = ['.eml', '.msg'];
    const fileName = file.name;
    const fileExtension = '.' + fileName.split('.').pop().toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('DropError', ['Only .eml or .msg files are supported']);
        return;
    }

    // Read the file content
    const reader = new FileReader();

    reader.onload = function(e) {
        const base64Content = btoa(
            new Uint8Array(e.target.result)
                .reduce((data, byte) => data + String.fromCharCode(byte), '')
        );

        // Parse email content to extract metadata
        const emailData = parseEmailFile(fileName, fileExtension.substring(1), e.target.result);

        // Send the parsed email data to AL
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod(
            'EmailParsed',
            [
                fileName,
                fileExtension.substring(1),
                base64Content,
                emailData.subject,
                emailData.senderEmail,
                emailData.senderName,
                emailData.receivedDate,
                emailData.hasAttachments
            ]
        );

        // If there are attachments, process each one
        if (emailData.attachments && emailData.attachments.length > 0) {
            emailData.attachments.forEach(function(attachment) {
                Microsoft.Dynamics.NAV.InvokeExtensibilityMethod(
                    'AttachmentParsed',
                    [
                        fileName,
                        attachment.fileName,
                        attachment.fileExtension,
                        attachment.mimeType,
                        attachment.content,
                        attachment.size
                    ]
                );
            });
        }

        // Signal that parsing is complete
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('EmailParsingComplete', []);
    };

    reader.onerror = function() {
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('DropError', ['Error reading the file']);
    };

    reader.readAsArrayBuffer(file);
}

function parseEmailFile(fileName, fileExtension, fileContent) {
    // This function would contain the actual email parsing logic
    // For demonstration purposes, we'll return mock data
    // In a real implementation, this would analyze the .eml or .msg file content

    const emailData = {
        subject: 'Sample Subject: ' + fileName,
        senderEmail: 'sender@example.com',
        senderName: 'John Doe',
        receivedDate: new Date().toISOString(),
        hasAttachments: true,
        attachments: []
    };

    // Add mock attachments
    if (emailData.hasAttachments) {
        // Mock attachment 1
        emailData.attachments.push({
            fileName: 'document1.pdf',
            fileExtension: 'pdf',
            mimeType: 'application/pdf',
            content: btoa('Mock PDF content'), // Base64 encoded content
            size: 1024
        });

        // Mock attachment 2
        emailData.attachments.push({
            fileName: 'spreadsheet.xlsx',
            fileExtension: 'xlsx',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            content: btoa('Mock Excel content'), // Base64 encoded content
            size: 2048
        });
    }

    return emailData;
}

function parseEmlFile(fileContent) {
    // In a real implementation, this would parse .eml format
    // .eml files are essentially text-based and can be parsed with regex or other text processing
    console.log("EML parsing would happen here");
}

function parseMsgFile(fileContent) {
    // In a real implementation, this would parse .msg format
    // .msg files are binary and would require specific libraries to parse
    console.log("MSG parsing would happen here");
}

function SetCustomerId(id) {
    customerId = id;
}

function GetControlHeight() {
    const dropZone = document.getElementById('dropZone');

    // Report a more compact height
    const height = dropZone ? Math.max(dropZone.offsetHeight, 120) : 120;

    console.log("Reporting control height to Business Central:", height);

    // Return the height via event
    try {
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlHeightReturned', [height]);
    } catch (e) {
        console.error("Error reporting height:", e);
    }
}

function SetPlaceholderText(text) {
    const messageElements = document.querySelectorAll('.drop-message');
    if (messageElements && messageElements.length > 0) {
        // Set the main drop message (first one)
        messageElements[0].textContent = text;

        // Make sure secondary message is still visible if it exists
        if (messageElements.length > 1) {
            messageElements[1].style.display = 'block';
        }
    }
}