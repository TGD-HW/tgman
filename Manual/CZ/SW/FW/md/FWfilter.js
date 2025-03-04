document.addEventListener("DOMContentLoaded", () => {
  const firmwareList = document.getElementById("firmware-list");
  const filters = {
    controlHW: document.getElementById("filter-controlHW"),
    NoAxis: document.getElementById("filter-NoAxis"),
    Comm: document.getElementById("filter-Comm"),
    FB: document.getElementById("filter-FB")
  };
  // Load firmware data
  fetch("FW.json") //works on local
  //fetch("/javascripts/FW.json")	//works on github
    .then(response => response.json())
    .then(data => {
      // Initial render
      renderList(data);

      // Attach event listeners for filtering
      Object.values(filters).forEach(filter => {
        filter.addEventListener("change", () => applyFilters(data));
      });
    })
    .catch(error => console.error("Error loading firmware data:", error));

  const applyFilters = (data) => {
    const filteredData = data.filter(item => {
      return (!filters.controlHW.value || item.controlHW === filters.controlHW.value) &&
             (!filters.NoAxis.value || item.NoAxis === filters.NoAxis.value) &&
             (!filters.Comm.value || item.Comm === filters.Comm.value) &&
             (!filters.FB.value || item.FB === filters.FB.value);
    });
    renderList(filteredData);
  };

  const renderList = (data) => {
    firmwareList.innerHTML = "";
    if (data.length === 0) {
      firmwareList.innerHTML = "<p>No firmware found matching the criteria.</p>";
      return;
    }
    data.forEach(item => {
      const div = document.createElement("div");
      div.innerHTML = `
        <p><strong>Control HW:</strong> ${item.controlHW}, <strong>Number of axis (type):</strong> ${item.NoAxis}, 
        <strong>Control communication type:</strong> ${item.Comm}, <strong>Feedback type:</strong> ${item.FB}</p>
        <p><a href="${item.FW}" target="_blank" rel="noopener noreferrer">Download Firmware</a></p>
      `;
      firmwareList.appendChild(div);
    });
  };
});