document.addEventListener("DOMContentLoaded", () => {
    const firmwareList = document.getElementById("firmware-list");
    const filters = {
        controlHW: document.getElementById("filter-controlHW"),
        NoAxis: document.getElementById("filter-NoAxis"),
        Comm: document.getElementById("filter-Comm"),
        FB: document.getElementById("filter-FB")
    };

    let firmwareData = []; // Store the fetched firmware data
    let firmwareCache = {}; // Cache for firmware files

    // Load firmware data
    fetch("FW.json") // works on local
        .then(response => response.json())
        .then(data => {
            firmwareData = data; // Store the data for filtering
            renderList(firmwareData); // Initial render
            // Attach event listeners for filtering
            Object.values(filters).forEach(filter => {
                filter.addEventListener("change", () => applyFilters());
            });
        })
        .catch(error => {
            console.error("Error loading firmware data:", error);
            firmwareList.innerHTML = "<p>Error loading firmware data. Please try again later.</p>";
        });

    const applyFilters = async () => {
        const filteredData = firmwareData.filter(item => {
            return (!filters.controlHW.value || item.controlHW === filters.controlHW.value) &&
                   (!filters.NoAxis.value || item.NoAxis === filters.NoAxis.value) &&
                   (!filters.Comm.value || item.Comm === filters.Comm.value) &&
                   (!filters.FB.value || item.FB === filters.FB.value);
        });
        await renderList(filteredData);
    };

	const getFirmwareFiles = async (folderPath) => {
    // Construct the URL for the specific folder
    const baseUrl = `https://firma1.tgdrives.cz:8443/!/#FW/view/head/${folderPath}`;

    const response = await fetch(baseUrl, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + btoa('FWpublic:FWpass1') // Encode credentials
        }
    });

    if (!response.ok) {
        console.error("Error fetching firmware files:", response.statusText);
        return [];
    }

    const text = await response.text(); // Get the response as text
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');

    // Extract file links from the HTML
    const fileLinks = Array.from(doc.querySelectorAll('a'))
        .map(link => link.href)
        .filter(href => href.endsWith('.bin') || href.endsWith('.zip')); // Adjust based on your file types

    return fileLinks;
};

    const renderList = async (data) => {
        firmwareList.innerHTML = "";
        if (data.length === 0) {
            firmwareList.innerHTML = "<p>No firmware found matching the criteria.</p>";
            return;
        }

        for (const item of data) {
            const div = document.createElement("div");
            div.innerHTML = `
                <p><strong>Control HW:</strong> ${item.controlHW}, <strong>Number of axis (type):</strong> ${item.NoAxis}, 
                <strong>Control communication type:</strong> ${item.Comm}, <strong>Feedback type:</strong> ${item.FB}</p>
                <p>Available Firmware Files:</p>
                <ul id="files-${item.FW}"></ul>
            `;
            firmwareList.appendChild(div);

            // Fetch and display firmware files
            const files = await getFirmwareFiles(item.FW);
            const fileList = document.getElementById(`files-${item.FW}`);
            if (files.length === 0) {
                fileList.innerHTML = "<li>No files found.</li>";
            } else {
                files.forEach(file => {
                    const listItem = document.createElement("li");
                    listItem.innerHTML = `<a href="${file}" target="_blank" rel="noopener noreferrer">${file.split('/').pop()}</a>`;
                    fileList.appendChild(listItem);
                });
            }
        }
    };
});
