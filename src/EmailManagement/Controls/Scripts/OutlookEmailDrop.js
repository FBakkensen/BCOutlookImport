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

    // Debounce the drag-over effect
    if (!this.classList.contains('drag-over')) {
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

    console.log('File dropped');

    // Check if files were dropped
    if (event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files.length > 0) {
        const file = event.dataTransfer.files[0];
        console.log('Processing file:', file.name);
        processFile(file);
    } else {
        console.warn('No files detected in drop event');
    }
}

function processFile(file) {
    console.log('Starting to process file:', file.name);

    // Check if file is an email (eml or msg)
    const validExtensions = ['.eml', '.msg'];
    const fileName = file.name;
    const fileExtension = '.' + fileName.split('.').pop().toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
        console.error('Unsupported file type:', fileExtension);
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('DropError', ['Only .eml or .msg files are supported']);
        return;
    }

    // Read the file content asynchronously
    const reader = new FileReader();

    reader.onload = function(e) {
        console.log('File read successfully:', fileName);

        // Get file content as string for parsing
        let fileContent;
        if (fileExtension === '.eml') {
            // For EML files, we can use text directly
            fileContent = new TextDecoder().decode(new Uint8Array(e.target.result));
        } else {
            // For other files, keep the binary data
            fileContent = e.target.result;
        }

        const base64Content = btoa(
            new Uint8Array(e.target.result)
                .reduce((data, byte) => data + String.fromCharCode(byte), '')
        );

        // Parse email content to extract metadata including body
        const emailData = parseEmailFile(fileName, fileExtension.substring(1), fileContent);
        console.log('Email data parsed:', emailData);

        // Log before sending to AL
        console.log('Invoking EmailParsed event with data:', {
            fileName,
            fileExtension: fileExtension.substring(1),
            base64Content,
            subject: emailData.subject,
            senderEmail: emailData.senderEmail,
            senderName: emailData.senderName,
            receivedDate: emailData.receivedDate,
            hasAttachments: emailData.hasAttachments,
            body: emailData.body
        });

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
                emailData.hasAttachments,
                emailData.body || '' // Include body, default to empty string if not available
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
        console.error('Error reading the file:', fileName);
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('DropError', ['Error reading the file']);
    };

    // Use appropriate read method based on file type
    if (fileExtension === '.eml') {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

function parseEmailFile(fileName, fileExtension, fileContent) {
    // Use the appropriate parser based on file extension
    let parserResult = {};
    let emailData = {
        subject: fileName, // Default to filename as subject
        senderEmail: '',
        senderName: '',
        receivedDate: new Date().toISOString(),
        hasAttachments: false,
        body: '',
        attachments: []
    };

    if (fileExtension === 'eml') {
        parserResult = parseEmlFile(fileContent);
    } else if (fileExtension === 'msg') {
        parserResult = parseMsgFile(fileContent);
    }

    // Merge parsed data with default values
    return { ...emailData, ...parserResult };
}

function parseEmlFile(fileContent) {
    console.log("==== STARTING EMAIL PARSE ====");
    console.log("Content length:", fileContent.length);

    // Basic EML parsing - extract headers and body
    let subject = '';
    let from = '';
    let senderEmail = '';
    let senderName = '';
    let receivedDate = new Date();
    let hasAttachments = false;
    let body = '';

    try {
        // Debug: Log a small sample of the content
        console.log("Email content sample:", fileContent.substring(0, 500));

        // Split the email into headers and body
        const headerBodySplit = fileContent.split(/\r?\n\r?\n/);
        const headers = headerBodySplit[0];

        // Log headers for debugging
        console.log("Headers:", headers);

        // Extract the raw body content (everything after headers)
        const rawBody = headerBodySplit.slice(1).join('\r\n\r\n');

        // Log a sample of the raw body
        console.log("Raw body sample:", rawBody.substring(0, 200));

        // Parse headers
        const subjectMatch = headers.match(/Subject: (.+?)(\r?\n|$)/i);
        if (subjectMatch) {
            subject = subjectMatch[1].trim();
            console.log("Subject found:", subject);
        }

        const fromMatch = headers.match(/From: (.+?)(\r?\n|$)/i);
        if (fromMatch) {
            from = fromMatch[1].trim();
            console.log("From found:", from);

            // Try to extract email and name from the From field
            const emailMatch = from.match(/<(.+?)>/);
            if (emailMatch) {
                senderEmail = emailMatch[1];
                // Extract name part (everything before the email)
                const namePart = from.split('<')[0].trim();
                senderName = namePart.replace(/"/g, '').replace(/'/g, '');
                console.log("Sender parsed:", senderName, senderEmail);
            } else {
                senderEmail = from;
                console.log("Sender email only:", senderEmail);
            }
        }

        // Check for content type and boundary in the headers
        console.log("Searching for Content-Type and boundary...");

        // Search for boundary markers in original form
        const boundaryRegex = /boundary="([^"]+)"/gi;
        const boundaryMatches = [...headers.matchAll(boundaryRegex)];
        console.log("Boundary matches found:", boundaryMatches.length);

        // Log all found boundary values
        boundaryMatches.forEach((match, index) => {
            console.log(`Boundary ${index + 1}:`, match[1]);
        });

        // Check for the MS Exchange specific pattern first
        const msExchangePattern = /boundary="_([0-9a-zA-Z]+)_"/i;
        const msExchangeMatch = headers.match(msExchangePattern);

        if (msExchangeMatch) {
            console.log("Microsoft Exchange boundary found:", msExchangeMatch[0]);
            console.log("Boundary value:", msExchangeMatch[1]);
        }

        // Handle multipart messages - first look for standard boundary definition
        let contentTypeMatch = headers.match(/Content-Type: multipart\/\w+;[\s\S]*?boundary="?([^"\r\n;]+)"?/i);

        // If not found, look for the special Microsoft Exchange boundary format
        if (!contentTypeMatch && msExchangeMatch) {
            contentTypeMatch = msExchangeMatch;
        }

        // Direct search for boundary attribute
        if (!contentTypeMatch) {
            const directBoundaryMatch = headers.match(/boundary="?([^"\r\n;]+)"?/i);
            if (directBoundaryMatch) {
                console.log("Direct boundary attribute found:", directBoundaryMatch[1]);
                contentTypeMatch = directBoundaryMatch;
            }
        }

        // As a fallback, look for any strings starting with --_ which are likely boundaries
        let boundary = '';
        if (contentTypeMatch) {
            boundary = contentTypeMatch[1];
            console.log("Boundary from content type match:", boundary);
        } else {
            // Look for potential boundary markers in the raw content
            console.log("Searching for boundary markers in content...");
            const boundaryMatch = rawBody.match(/--_[0-9a-zA-Z]+_/);
            if (boundaryMatch) {
                boundary = boundaryMatch[0].substring(2); // Remove the leading --
                console.log("Found boundary in content:", boundary);
            } else {
                // Last resort: search for anything that looks like a boundary
                const genericBoundaryMatch = rawBody.match(/--([A-Za-z0-9._=]+)/);
                if (genericBoundaryMatch) {
                    boundary = genericBoundaryMatch[1];
                    console.log("Found generic boundary:", boundary);
                }
            }
        }

        // Remove the special keyword "boundary=" if it's part of the detected boundary
        boundary = boundary.replace(/^boundary=/i, '');

        // Remove quotes if present
        boundary = boundary.replace(/^"(.*)"$/, '$1');

        console.log("Final boundary value:", boundary);

        // Special handling for the Microsoft-style boundaries
        if (boundary.startsWith('_') && boundary.endsWith('_')) {
            // This is the exact format we're seeing with
            // "boundary="_000_AS8P191MB1846C001057EC39E8D1FE923B9CA2AS8P191MB1846EURP_""
            console.log("Detected Microsoft-style boundary format");

            // Extract all text blocks that are likely to be content
            const contentBlocks = rawBody.split(/--_[0-9a-zA-Z]+_/).filter(part => part.trim().length > 0);
            console.log("Content blocks found:", contentBlocks.length);

            // Process each content block
            let textContent = '';
            let htmlContent = '';

            for (let i = 0; i < contentBlocks.length; i++) {
                const block = contentBlocks[i];
                console.log(`Analyzing block ${i+1}, length: ${block.length}`);

                // Check if this is plain text content
                if (block.includes('Content-Type: text/plain')) {
                    const parts = block.split(/\r?\n\r?\n/);
                    if (parts.length > 1) {
                        textContent = parts.slice(1).join('\r\n\r\n');
                        console.log("Found text/plain content, length:", textContent.length);

                        // Check if this content is Base64 encoded
                        if (block.includes('Content-Transfer-Encoding: base64') ||
                            textContent.match(/^[A-Za-z0-9+/=\s]+$/)) {
                            console.log("Detected Base64 encoded text content, attempting to decode");
                            try {
                                // Remove whitespace before decoding
                                const cleanBase64 = textContent.replace(/\s/g, '');
                                if (cleanBase64.match(/^[A-Za-z0-9+/=]+$/)) {
                                    const decodedText = atob(cleanBase64);
                                    console.log("Successfully decoded Base64 text content, length:", decodedText.length);
                                    textContent = decodedText;
                                }
                            } catch (error) {
                                console.error("Failed to decode Base64 text content:", error);
                            }
                        }
                    }
                }

                // Check if this is HTML content
                if (block.includes('Content-Type: text/html')) {
                    const parts = block.split(/\r?\n\r?\n/);
                    if (parts.length > 1) {
                        htmlContent = parts.slice(1).join('\r\n\r\n');
                        console.log("Found text/html content, length:", htmlContent.length);

                        // Check if this content is Base64 encoded
                        if (block.includes('Content-Transfer-Encoding: base64') ||
                            htmlContent.match(/^[A-Za-z0-9+/=\s]+$/)) {
                            console.log("Detected Base64 encoded HTML content, attempting to decode");
                            try {
                                // Remove whitespace before decoding
                                const cleanBase64 = htmlContent.replace(/\s/g, '');
                                if (cleanBase64.match(/^[A-Za-z0-9+/=]+$/)) {
                                    const decodedHtml = atob(cleanBase64);
                                    console.log("Successfully decoded Base64 HTML content, length:", decodedHtml.length);
                                    htmlContent = decodedHtml;
                                }
                            } catch (error) {
                                console.error("Failed to decode Base64 HTML content:", error);
                            }
                        }
                    }
                }
            }

            // Use text content if available, otherwise use HTML with tags stripped
            if (textContent) {
                body = textContent;
                console.log("Using text content for body");
            } else if (htmlContent) {
                body = htmlContent.replace(/<[^>]*>/g, ' ')
                    .replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&#39;/g, "'");
                console.log("Using HTML content (stripped) for body");
            } else {
                // Fallback to the largest content block
                const largestBlock = contentBlocks.reduce((largest, current) =>
                    current.length > largest.length ? current : largest, '');

                // Strip headers from the block
                body = largestBlock.replace(/Content-Type:.*?\r?\n\r?\n/is, '')
                    .replace(/Content-Transfer-Encoding:.*?\r?\n/ig, '');
                console.log("Using largest content block as fallback, length:", body.length);
            }
        } else if (boundary) {
            // This is a multipart message - extract parts using the boundary
            console.log("Processing standard multipart message");

            // Prepare safe boundary pattern for regex - both with and without leading --
            const boundaryPattern = new RegExp(`(--|)${boundary.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`, 'g');
            console.log("Using boundary pattern:", boundaryPattern);

            // Split by boundary
            const parts = rawBody.split(boundaryPattern).filter(part => part.trim() && !part.includes('--'));
            console.log("Found parts:", parts.length);

            // Debug: log each part
            parts.forEach((part, index) => {
                console.log(`Part ${index+1} sample:`, part.substring(0, 100));
            });

            // Find the text/plain or text/html part
            let textBody = '';
            let htmlBody = '';

            for (const part of parts) {
                if (part.trim().length === 0) continue;

                // Check if this part is text/plain
                if (part.match(/Content-Type:\s*text\/plain/i)) {
                    // Extract content after the headers in this part
                    const partContent = part.split(/\r?\n\r?\n/).slice(1).join('\r\n\r\n');
                    textBody = partContent.trim();
                    console.log("Found text/plain part, length:", textBody.length);
                }

                // Check if this part is text/html
                if (part.match(/Content-Type:\s*text\/html/i)) {
                    // Extract content after the headers in this part
                    const partContent = part.split(/\r?\n\r?\n/).slice(1).join('\r\n\r\n');
                    htmlBody = partContent.trim();
                    console.log("Found text/html part, length:", htmlBody.length);
                }
            }

            // Prefer text body, but use HTML if that's all we have
            if (textBody) {
                body = textBody;
                console.log("Using text body for content");
            } else if (htmlBody) {
                // Strip HTML tags for readability
                body = htmlBody.replace(/<[^>]*>/g, ' ')
                    .replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&#39;/g, "'");
                console.log("Using HTML body for content (with tags stripped)");
            } else {
                console.log("Could not find text or HTML parts, trying aggressive extraction");
                // If we couldn't find a clear text or HTML part, extract content more aggressively
                // Look for content between header sections and boundary markers
                const contentMatches = rawBody.match(/(?:Content-Type:.*?\r?\n\r?\n)([\s\S]*?)(?=\r?\n\r?\n--)/g);
                if (contentMatches && contentMatches.length > 0) {
                    console.log("Found content matches:", contentMatches.length);
                    // Use the longest content section as the body
                    body = contentMatches.reduce((longest, current) =>
                        current.length > longest.length ? current : longest, '');

                    // Remove headers
                    body = body.replace(/Content-Type:.*?\r?\n\r?\n/i, '');
                    console.log("Using longest content match, length:", body.length);
                } else {
                    // Last resort: just use everything after removing boundary lines
                    console.log("Using last resort body extraction");
                    body = rawBody.replace(new RegExp(`.*?${boundary}.*?\r?\n`, 'g'), '\n')
                        .replace(/Content-Type:.+?\r?\n/g, '')
                        .replace(/Content-Transfer-Encoding:.+?\r?\n/g, '')
                        .trim();
                    console.log("Last resort body length:", body.length);
                }
            }
        } else {
            // Not a multipart message, just use the body directly
            console.log("Not a multipart message, using raw body");
            body = rawBody;

            // If it's HTML, strip the tags
            if (headers.includes('Content-Type: text/html')) {
                console.log("Stripping HTML tags from body");
                body = body.replace(/<[^>]*>/g, ' ')
                    .replace(/&amp;/g, '&')
                    .replace(/&lt;/g, '<')
                    .replace(/&gt;/g, '>')
                    .replace(/&quot;/g, '"')
                    .replace(/&#39;/g, "'");
            }
        }

        // Clean up any quoted-printable encoding
        if (headers.includes('Content-Transfer-Encoding: quoted-printable')) {
            console.log("Decoding quoted-printable content");
            body = body.replace(/=\r?\n/g, '')  // Soft line breaks
                .replace(/=([0-9A-F]{2})/gi, (match, p1) => String.fromCharCode(parseInt(p1, 16)));
        }

        // Decode Base64 encoded content
        if (headers.includes('Content-Transfer-Encoding: base64') || body.match(/^[A-Za-z0-9+/=\s]+$/)) {
            console.log("Detected Base64 encoded content, decoding...");
            try {
                // Remove any whitespace before decoding
                const cleanBase64 = body.replace(/\s/g, '');

                // Check if it's valid Base64
                if (cleanBase64.match(/^[A-Za-z0-9+/=]+$/)) {
                    const decodedBody = atob(cleanBase64);
                    console.log("Base64 decoded body length:", decodedBody.length);
                    body = decodedBody;
                } else {
                    console.log("Content looked like Base64 but failed validation check");
                }
            } catch (error) {
                console.error("Failed to decode Base64 content:", error);
            }
        }

        // Clean up extra whitespace and boundary markers
        console.log("Body before final cleanup:", body.substring(0, 100));

        // Remove any text that looks like a boundary marker or the word "boundary"
        body = body.replace(/\r\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')  // Limit consecutive newlines
            .replace(/--_[0-9a-zA-Z]+_/g, '') // Remove Microsoft-style boundary markers
            .replace(/--\[.*?\]--/g, '')  // Remove other boundary formats
            .replace(/--.*?--/g, '')     // Remove other boundary formats
            .replace(/boundary="_[0-9a-zA-Z]+_"/gi, '') // Remove boundary attributes
            .replace(/boundary=/gi, '')  // Remove boundary keyword
            .replace(/_[0-9a-zA-Z]+_/g, '') // Remove underscore-wrapped markers
            .trim();

        console.log("Final body sample:", body.substring(0, 100));
        console.log("Final body length:", body.length);
        console.log("==== EMAIL PARSE COMPLETE ====");
    } catch (e) {
        console.error("Error parsing EML file:", e);
        body = "Error parsing email content: " + e.message;
    }

    return {
        subject,
        senderEmail,
        senderName,
        receivedDate: receivedDate.toISOString(),
        hasAttachments,
        body
    };
}

function parseMsgFile(fileContent) {
    console.log("Attempting to parse MSG file content");

    // MSG parsing is complex and requires specialized libraries
    // For now, we'll do a basic extraction of text content

    let body = '';

    try {
        // Convert binary data to a string and look for text patterns
        const rawText = new TextDecoder('utf-8').decode(fileContent);

        // Try to find blocks of readable text (very basic approach)
        const bodyMatch = rawText.match(/[\x20-\x7E\s]{100,}/g);
        if (bodyMatch) {
            // Use the longest text block as the body
            body = bodyMatch.reduce((longest, current) =>
                current.length > longest.length ? current : longest, '');
        } else {
            body = "This MSG file could not be parsed properly. Please use the 'Download Email File' option to view the original.";
        }
    } catch (e) {
        console.error("Error parsing MSG file:", e);
        body = "Error parsing MSG file content: " + e.message;
    }

    return {
        // We can't reliably extract other metadata from MSG without a proper library
        body: body
    };
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