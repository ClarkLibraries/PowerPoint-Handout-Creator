// script.js

// IMPORTANT: These global variables (VFLauncher, tinycolor)
// are assumed to be exposed by the .min.js files loaded in index.html
// You MUST ensure the correct .min.js files for @vf.js/launcher and tinycolor2
// are included in index.html, and that they expose these global names.

// Access pptx-parser's functions via the global 'pptx-parser' object
// The UMD build you provided should make the main 'parse' function available
// as the 'default' export, and 'vf' as a named export on that object.
const pptxParserGlobal = window['pptx-parser']; // Get the entire object
const parsePptx = pptxParserGlobal ? pptxParserGlobal.default : null; // Access the 'default' export
const vfConvert = pptxParserGlobal ? pptxParserGlobal.vf : null; // Access the 'vf' named export

// Hypothetical global names for other libraries.
// YOU NEED TO VERIFY THESE based on their actual UMD builds.
const createVF = window.VFLauncher ? window.VFLauncher.createVF : null; // e.g., if VFLauncher exposes createVF
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
    const parseButton = document.getElementById('parseButton');
    const prevButton = document.getElementById('btn-prev');
    const nextButton = document.getElementById('btn-next');
    const outputDiv = document.getElementById('output');
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    const vfContainer = document.querySelector('.vf-container');

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

    // Initial state setup
    parseButton.disabled = true;
    prevButton.disabled = true;
    nextButton.disabled = true;

    // Enable parse button once a file is chosen
    fileDlg.addEventListener('change', () => {
        if (fileDlg.files.length > 0) {
            parseButton.disabled = false;
            outputDiv.innerHTML = ''; // Clear previous output
            errorMessage.style.display = 'none';
        } else {
            parseButton.disabled = true;
        }
    });

    // Check if required global libraries are available
    if (!parsePptx || !vfConvert || !createVF || !tinycolor2) {
        errorMessage.textContent = 'Required JavaScript libraries (pptx-parser, @vf.js/launcher, tinycolor2) are not loaded correctly. Please ensure all .min.js files are in the lib/ folder and linked in index.html.';
        errorMessage.style.display = 'block';
        console.error('Missing global library functions:', {
            pptxParserGlobal: !!pptxParserGlobal, // Check the main object
            parsePptx: !!parsePptx,
            vfConvert: !!vfConvert,
            createVF: !!createVF,
            tinycolor2: !!tinycolor2
        });
    }

});

function handleFileSelection(e) {
    console.log("File selected:", e.target.files[0] ? e.target.files[0].name : "No file");
}

async function handleParseClick() {
    const file = document.getElementById('pptxFile').files[0];
    const parseButton = document.getElementById('parseButton');
    const outputDiv = document.getElementById('output');
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    const prevButton = document.getElementById('btn-prev');
    const nextButton = document.getElementById('btn-next');
    const vfContainer = document.querySelector('.vf-container');

    outputDiv.innerHTML = '';
    errorMessage.style.display = 'none';

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
    prevButton.disabled = true;
    nextButton.disabled = true;

    try {
        console.log("Starting PPTX parsing...");
        // Use the corrected 'parsePptx' function
        const pptJson = await parsePptx(file, { flattenGroup: true });
        console.log("Parsed PPTX JSON:", pptJson);

        if (!pptJson || !pptJson.pageSize) {
            throw new Error("Failed to extract page size from PPTX. Invalid PPTX structure or parsing error.");
        }

        const width = pptJson.pageSize.width.value;
        const height = pptJson.pageSize.height.value;

        if (vfConvert) {
            console.log("Converting to VF JSON...");
            // Use the corrected 'vfConvert' function
            const vfJson = await vfConvert(pptJson, { width, height });
            console.log('VF JSON:', vfJson);

            const tmp = new Blob([JSON.stringify(vfJson)], { type: 'application/json' });

            const config = {
                container: vfContainer,
                debug: true,
                width,
                height,
                resolution: window.devicePixelRatio || 1
            };
            console.log("VF Player Config:", config);

            if (createVF) {
                const v = createVF(config, player => {
                    window.player = playerInstance = player;
                    window.v = v;

                    player.onReady = function() {
                        console.log("VF Player: onReady");
                        currentSceneIndex = 0;
                        totalScenes = playerInstance.data.scenes.length;
                        if (totalScenes > 0) {
                            prevButton.disabled = false;
                            nextButton.disabled = false;
                        }
                        displayHandoutContent(playerInstance.data.scenes);
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
                 console.error('createVF is null. Check @vf.js/launcher.min.js.');
            }
        } else {
            errorMessage.textContent = 'PPTX to VF conversion function (vf) not found. Cannot render presentation.';
            errorMessage.style.display = 'block';
            console.error('vfConvert is null. Check pptx-parser.min.js for vf export.');
            displayHandoutContentFromPptJson(pptJson);
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

function displayHandoutContent(scenes) {
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
