// script.js

// IMPORTANT: These global variables (VFLauncher, tinycolor)
// are assumed to be exposed by the .min.js files loaded in index.html
// You MUST ensure the correct .min.js files for @vf.js/launcher and tinycolor2
// are included in index.html, and that they expose these global names.

// Access pptx-parser's functions via the global 'pptx-parser' object
const pptxParserGlobal = window['pptx-parser']; // Get the entire object
const parsePptx = pptxParserGlobal ? pptxParserGlobal.default : null; // Access the 'default' export
const vfConvert = pptxParserGlobal ? pptxParserGlobal.vf : null; // Access the 'vf' named export

// CORRECTED: createVF is directly on the window object
const createVF = window.createVF; // This is the corrected line!
const tinycolor2 = window.tinycolor; // Common global name for tinycolor2

// Exposing tinycolor2 globally as 'tiny' as per your original snippet
if (tinycolor2) {
    window.tiny = tinycolor2;
} else {
    console.warn("tinycolor2 global not found. 'window.tiny' will not be set.");
}

let playerInstance = null;
let totalScenes = 0;
let currentSceneIndex = 0;

document.addEventListener('DOMContentLoaded', () => {
    const fileDlg = document.getElementById('pptxFile');
    const fileLabel = document.getElementById('fileLabel'); // Get the label element
    const parseButton = document.getElementById('parseButton');
    const downloadHandoutBtn = document.getElementById('downloadHandoutBtn'); // Get the download button
    const clearBtn = document.getElementById('clearBtn'); // Get the clear button
    const prevButton = document.getElementById('btn-prev');
    const nextButton = document.getElementById('btn-next');
    const outputDiv = document.getElementById('output');
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    const vfPlayerContainer = document.getElementById('vfPlayerContainer'); // Get the VF player container
    const processStatus = document.getElementById('process-status');
    const handoutPlaceholder = document.getElementById('handoutPlaceholder');

    fileDlg.onchange = handleFileSelection;
    parseButton.onclick = handleParseClick;

    prevButton.onclick = function() {
        if (playerInstance && currentSceneIndex > 0) {
            currentSceneIndex--;
            playerInstance.switchToSceneIndex(currentSceneIndex);
            console.log(`Switched to scene: ${currentSceneIndex + 1}`);
        }
    };
    nextButton.onclick = function() {
        if (playerInstance && currentSceneIndex < totalScenes - 1) {
            currentSceneIndex++;
            playerInstance.switchToSceneIndex(currentSceneIndex);
            console.log(`Switched to scene: ${currentSceneIndex + 1}`);
        }
    };

    // Clear All button functionality
    clearBtn.onclick = function() {
        fileDlg.value = null; // Clear the file input
        fileLabel.classList.remove('has-files'); // Remove visual feedback
        processStatus.textContent = 'No file selected.';
        outputDiv.innerHTML = '';
        handoutPlaceholder.style.display = 'block'; // Show handout placeholder
        errorMessage.style.display = 'none';
        loadingMessage.style.display = 'none';
        parseButton.disabled = true;
        downloadHandoutBtn.disabled = true;
        prevButton.disabled = true;
        nextButton.disabled = true;

        // Clear VF Player
        if (playerInstance) {
            playerInstance.dispose(); // Dispose of the VF player instance
            playerInstance = null;
        }
        vfPlayerContainer.innerHTML = '<p>Upload a .pptx file to see the live presentation preview.</p>'; // Reset VF player content
        totalScenes = 0;
        currentSceneIndex = 0;
    };


    // Initial state setup
    parseButton.disabled = true;
    downloadHandoutBtn.disabled = true; // Initially disable download
    prevButton.disabled = true;
    nextButton.disabled = true;

    // Enable parse button once a file is chosen
    fileDlg.addEventListener('change', () => {
        if (fileDlg.files.length > 0) {
            parseButton.disabled = false;
            processStatus.textContent = `File selected: ${fileDlg.files[0].name}`;
            fileLabel.classList.add('has-files'); // Add visual feedback
            outputDiv.innerHTML = ''; // Clear previous output
            handoutPlaceholder.style.display = 'block'; // Show handout placeholder initially
            errorMessage.style.display = 'none';
            loadingMessage.style.display = 'none';
            // Reset VF Player content when a new file is chosen
            if (playerInstance) {
                playerInstance.dispose();
                playerInstance = null;
            }
            vfPlayerContainer.innerHTML = '<p>Upload a .pptx file to see the live presentation preview.</p>';
        } else {
            parseButton.disabled = true;
            processStatus.textContent = 'No file selected.';
            fileLabel.classList.remove('has-files'); // Remove visual feedback
        }
    });

    // Drag and drop functionality for the file upload area
    fileLabel.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileLabel.classList.add('border-blue-600', 'bg-blue-50'); // Visual feedback on drag over
    });

    fileLabel.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileLabel.classList.remove('border-blue-600', 'bg-blue-50'); // Remove feedback on drag leave
    });

    fileLabel.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileLabel.classList.remove('border-blue-600', 'bg-blue-50');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            // Assign the dropped file to the input
            fileDlg.files = files;
            // Manually trigger the change event to update UI
            const event = new Event('change', { bubbles: true });
            fileDlg.dispatchEvent(event);
        }
    });

    // Check if required global libraries are available
    if (!pptxParserGlobal || !parsePptx || !vfConvert || !createVF || !tinycolor2) {
        errorMessage.textContent = 'Required JavaScript libraries (pptx-parser, @vf.js/launcher, tinycolor2) are not loaded correctly. Please ensure all .min.js files are in the lib/ folder and linked in index.html.';
        errorMessage.style.display = 'block';
        parseButton.disabled = true;
        console.error('Missing global library functions:', {
            pptxParserGlobal: !!pptxParserGlobal,
            parsePptx: !!parsePptx,
            vfConvert: !!vfConvert,
            createVF: !!createVF, // This will now correctly reflect if window.createVF exists
            tinycolor2: !!tinycolor2
        });
    }

    // Placeholder for download functionality (you'll implement this)
    downloadHandoutBtn.addEventListener('click', () => {
        // Logic to generate and download the HTML handout
        console.log('Download Handout button clicked!');
        alert('Download Handout functionality coming soon!');
    });

});

function handleFileSelection(e) {
    console.log("File selected:", e.target.files[0] ? e.target.files[0].name : "No file");
}

async function handleParseClick() {
    const file = document.getElementById('pptxFile').files[0];
    const parseButton = document.getElementById('parseButton');
    const downloadHandoutBtn = document.getElementById('downloadHandoutBtn');
    const outputDiv = document.getElementById('output');
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    const prevButton = document.getElementById('btn-prev');
    const nextButton = document.getElementById('btn-next');
    const vfPlayerContainer = document.getElementById('vfPlayerContainer');
    const handoutPlaceholder = document.getElementById('handoutPlaceholder');

    // --- ADD THESE LOGS ---
    console.log('--- Debugging handleParseClick ---');
    console.log('file:', file);
    console.log('parseButton:', parseButton);
    console.log('downloadHandoutBtn:', downloadHandoutBtn);
    console.log('outputDiv:', outputDiv);
    console.log('loadingMessage:', loadingMessage);
    console.log('errorMessage:', errorMessage);
    console.log('prevButton:', prevButton);
    console.log('nextButton:', nextButton);
    console.log('vfPlayerContainer:', vfPlayerContainer);
    console.log('handoutPlaceholder:', handoutPlaceholder);
    console.log('---------------------------------');
    // --- END ADDED LOGS ---


    outputDiv.innerHTML = ''; // Clear previous output
    handoutPlaceholder.style.display = 'none'; // Hide placeholder
    errorMessage.style.display = 'none'; // <--- This is the line where the error occurs
    if (!file) {
        errorMessage.textContent = 'Please select a .pptx file first.';
        errorMessage.style.display = 'block';
        return;
    }

    // Double-check if the functions are available after the DOM is ready
    if (!parsePptx || !vfConvert || !createVF) {
        errorMessage.textContent = 'Critical parsing/rendering libraries are not available. Check browser console for details.';
        errorMessage.style.display = 'block';
        return;
    }

    loadingMessage.style.display = 'block';
    parseButton.disabled = true;
    downloadHandoutBtn.disabled = true;
    prevButton.disabled = true;
    nextButton.disabled = true;

    // Dispose existing player if any
    if (playerInstance) {
        playerInstance.dispose();
        playerInstance = null;
    }
    vfPlayerContainer.innerHTML = ''; // Clear VF player content area

    try {
        console.log("Starting PPTX parsing...");
        const pptJson = await parsePptx(file, { flattenGroup: true });
        console.log("Parsed PPTX JSON:", pptJson);

        if (!pptJson || !pptJson.pageSize) {
            throw new Error("Failed to extract page size from PPTX. Invalid PPTX structure or parsing error.");
        }

        const width = pptJson.pageSize.width.value;
        const height = pptJson.pageSize.height.value;

        if (vfConvert) {
            console.log("Converting to VF JSON...");
            const vfJson = await vfConvert(pptJson, { width, height });
            console.log('VF JSON:', vfJson);

            const tmp = new Blob([JSON.stringify(vfJson)], { type: 'application/json' });

            const config = {
                container: vfPlayerContainer, // Use the dedicated container
                debug: true,
                width,
                height,
                resolution: window.devicePixelRatio || 1 // Fallback for resolution
            };
            console.log("VF Player Config:", config);

            if (createVF) {
                // The createVF function itself is the global
                const v = createVF(config, player => { // No longer window.VFLauncher.createVF
                    window.player = playerInstance = player;
                    window.v = v; // Store the VF instance globally

                    player.onReady = function() {
                        console.log("VF Player: onReady"); // Initialization complete
                        currentSceneIndex = 0;
                        totalScenes = playerInstance.data.scenes.length;
                        if (totalScenes > 0) {
                            prevButton.disabled = false;
                            nextButton.disabled = false;
                        }
                        displayHandoutContent(playerInstance.data.scenes, pptJson); // Pass pptJson for notes
                        downloadHandoutBtn.disabled = false; // Enable download button
                    };

                    player.onSceneCreate = function() {
                        console.log("VF Player: onSceneCreate");
                    };

                    player.onMessage = function(msg) {
                        console.log("VF Player: onMessage ==>", msg);
                    };

                    player.onError = function(evt) {
                        console.error("VF Player: onError ==>", evt);
                        errorMessage.textContent = `VF Player Error: ${evt.message || evt}`;
                        errorMessage.style.display = 'block';
                    };

                    player.onDispose = function() {
                        console.log("VF Player: onDispose");
                        playerInstance = null;
                        prevButton.disabled = true;
                        nextButton.disabled = true;
                    };

                    player.play(URL.createObjectURL(tmp));
                });
            } else {
                 errorMessage.textContent = 'VF.js launcher (createVF) not found. Cannot render presentation.';
                 errorMessage.style.display = 'block';
                 console.error('createVF is null. Check vf-launcher.min.js.');
                 displayHandoutContentFromPptJson(pptJson); // Still display text handout
            }
        } else {
            errorMessage.textContent = 'PPTX to VF conversion function (vf) not found. Cannot render presentation.';
            errorMessage.style.display = 'block';
            console.error('vfConvert is null. Check pptx-parser.min.js for vf export.');
            displayHandoutContentFromPptJson(pptJson); // Fallback for text-only handout if vf is not available
            downloadHandoutBtn.disabled = false; // Still allow download for text handout
        }

    } catch (e) {
        console.error('An error occurred during parsing or rendering:', e);
        errorMessage.textContent = `Error: ${e.message || 'An unknown error occurred.'}`;
        errorMessage.style.display = 'block';
    } finally {
        loadingMessage.style.display = 'none';
        parseButton.disabled = false;
    }
}

// Pass pptJson to this function to access speaker notes
function displayHandoutContent(scenes, pptJson) {
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = '<h2>Handout Content:</h2>';

    if (!scenes || scenes.length === 0) {
        outputDiv.innerHTML += '<p>No content extracted for handout.</p>';
        return;
    }

    scenes.forEach((scene, index) => {
        const slideDiv = document.createElement('div');
        slideDiv.classList.add('slide-content');

        const slideTitle = document.createElement('h3');
        slideTitle.textContent = `Slide ${index + 1}`;

        let slideText = [];
        if (scene.elementMap) {
            for (const key in scene.elementMap) {
                const element = scene.elementMap[key];
                if (element.text) {
                    slideText.push(element.text);
                }
            }
        }

        if (slideText.length > 0) {
            slideTitle.textContent += ': ' + slideText[0].substring(0, Math.min(slideText[0].length, 50)) + '...';
            slideText.forEach(text => {
                const p = document.createElement('p');
                p.textContent = text;
                slideDiv.appendChild(p);
            });
        } else {
            const p = document.createElement('p');
            p.textContent = 'No text content directly extractable from VF scene data for this slide.';
            slideDiv.appendChild(p);
        }

        // Extract and display speaker notes from the original pptJson
        const originalSlide = pptJson.slides[index];
        if (originalSlide && originalSlide.notes && originalSlide.notes.text) {
            const notesP = document.createElement('p');
            notesP.innerHTML = `<strong>Speaker Notes:</strong> ${originalSlide.notes.text}`;
            slideDiv.appendChild(notesP);
        } else {
            const p = document.createElement('p');
            p.textContent = 'No speaker notes found for this slide.';
            slideDiv.appendChild(p);
        }

        slideDiv.prepend(slideTitle);
        outputDiv.appendChild(slideDiv);
    });
}

function displayHandoutContentFromPptJson(pptJson) {
    const outputDiv = document.getElementById('output');
    outputDiv.innerHTML = '<h2>Handout Content (Text-Only Fallback):</h2>';

    if (!pptJson || !Array.isArray(pptJson.slides)) {
        outputDiv.innerHTML += '<p>Could not parse slides from the PPTX file.</p>';
        return;
    }

    pptJson.slides.forEach((slide, index) => {
        const slideDiv = document.createElement('div');
        slideDiv.classList.add('slide-content');

        const slideTitle = document.createElement('h3');
        slideTitle.textContent = `Slide ${index + 1}`;
        slideDiv.appendChild(slideTitle);

        if (slide.text && Array.isArray(slide.text)) {
            slide.text.forEach(textBlock => {
                const p = document.createElement('p');
                p.textContent = textBlock;
                slideDiv.appendChild(p);
            });
        } else if (slide.text && typeof slide.text === 'string') {
             const p = document.createElement('p');
             p.textContent = slide.text;
             slideDiv.appendChild(p);
        } else {
            const p = document.createElement('p');
            p.textContent = 'No main slide text found.';
            slideDiv.appendChild(p);
        }

        if (slide.notes && slide.notes.text) {
            const notesP = document.createElement('p');
            notesP.innerHTML = `<strong>Speaker Notes:</strong> ${slide.notes.text}`;
            slideDiv.appendChild(notesP);
        } else {
            const p = document.createElement('p');
            p.textContent = 'No speaker notes found for this slide.';
            slideDiv.appendChild(p);
        }

        outputDiv.appendChild(slideDiv);
    });
}
