Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        // Initialize the task pane
        fetchSVGImagesFromGitHub();
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "block";

        // Add event listener for search button
        document.getElementById('search-button').addEventListener('click', searchSVGs);

        // Add event listener for dynamic search input
        document.getElementById('search-input').addEventListener('input', searchSVGs);
    }
});

async function fetchSVGImagesFromGitHub() {
    try {
        // GitHub API endpoint to fetch files from the BioDraws repository
        const response = await fetch('https://api.github.com/repos/bioDraws/BioDRaws/contents');
        const files = await response.json();

        const svgContainer = document.getElementById('svg-container');
        svgContainer.innerHTML = ''; // Clear existing content

        // Store the files globally to use them for searching
        window.svgFiles = files.filter(file => file.name.endsWith('.svg'));

        // Display all SVG files initially
        displaySVGThumbnails(window.svgFiles);
    } catch (error) {
        console.error('Error fetching SVG images from GitHub:', error);
    }
}

function displaySVGThumbnails(files: any[]) {
    const svgContainer = document.getElementById('svg-container');
    svgContainer.innerHTML = ''; // Clear existing content

    files.forEach(file => {
        const thumbnailLabel = document.createElement('div');
        thumbnailLabel.className = 'svg-thumbnail-container';

        const img = document.createElement('img');
        img.src = `https://raw.githubusercontent.com/bioDraws/BioDRaws/main/${file.path}`;
        img.className = 'svg-thumbnail';
        img.onclick = () => importSVGToPowerPoint(img.src); // When clicked, import into PowerPoint

        // Remove 'biodraw_' or 'biodraws_' at the beginning and '.svg' at the end
        const fileName = file.name.replace(/^biodraws?_/, '').replace(/\.svg$/, '');

        const nameLabel = document.createElement('div');
        nameLabel.className = 'svg-thumbnail-name';
        nameLabel.textContent = fileName;

        thumbnailLabel.appendChild(img);
        thumbnailLabel.appendChild(nameLabel);
        svgContainer.appendChild(thumbnailLabel);
    });
}

function importSVGToPowerPoint(svgUrl: string) {
    fetch(svgUrl)
        .then(response => response.text())
        .then(svgContent => {
            // Use the Office JavaScript API to insert the SVG
            Office.context.document.setSelectedDataAsync(svgContent, {
                coercionType: Office.CoercionType.XmlSvg
            }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to insert SVG:', asyncResult.error.message);
                } else {
                    console.log('SVG inserted successfully');
                }
            });
        })
        .catch(error => {
            console.error('Error importing SVG:', error);
        });
}

// Function to filter SVGs based on the search input
function searchSVGs() {
    const searchQuery = document.getElementById('search-input').value.toLowerCase();
    const filteredFiles = window.svgFiles.filter(file => file.name.toLowerCase().includes(searchQuery));
    displaySVGThumbnails(filteredFiles);
}
