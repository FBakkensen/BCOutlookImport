var customerId = '';

function InitializeControl() {
    const dropZone = document.createElement('div');
    dropZone.id = 'dropZone';
    dropZone.className = 'outlook-drop-zone';
    dropZone.innerHTML = '<div class="drop-message">Drop Outlook email here</div>';
    
    // Setup drag-and-drop event listeners
    dropZone.addEventListener('dragover', handleDragOver);
    dropZone.addEventListener('dragleave', handleDragLeave);
    dropZone.addEventListener('drop', handleFileDrop);
    dropZone.addEventListener('click', handleClick);
    
    document.body.appendChild(dropZone);
    
    // Notify AL that the control is ready
    Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlReady', []);
}

function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    this.classList.add('drag-over');
}

function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    this.classList.remove('drag-over');
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
    const height = dropZone ? dropZone.offsetHeight : 150;
    
    // Return the height via event instead of direct return
    Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('ControlHeightReturned', [height]);
}

function SetPlaceholderText(text) {
    const messageDiv = document.querySelector('.drop-message');
    if (messageDiv) {
        messageDiv.textContent = text;
    }
}