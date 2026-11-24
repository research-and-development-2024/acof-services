(function () {
    "use strict";

    document.addEventListener("DOMContentLoaded", function () {
        loadTenantResources();
    });

    async function loadTenantResources() {
        try {
            const resourcesUrl = "data/resources.json?v=" + Date.now();

            const response = await fetch(resourcesUrl, {
                cache: "no-store"
            });

            if (!response.ok) {
                throw new Error("Could not load resources data (resources.json)");
            }

            const spaData = await response.json();
            renderSPASections(spaData);
        }

        catch (error) {
            console.error("Error loading tenant resources:", error);
        }
    }

    function renderSPASections(spaData) {
        const container = document.getElementById("spa-sections-container");
        const template = document.getElementById("spa-section-template");

        spaData.forEach(function (spa) {
            if (!spa.buildings || spa.buildings.length === 0) {
                return;
            }

            const section = template.content.cloneNode(true);

            section.querySelector('[data-field="spa-display-name"]').textContent = spa["display-name"];

            const buildingsList = section.querySelector('[data-field="buildings-list"]');

            spa.buildings.forEach(function (building) {
                const li = document.createElement("li");
                li.textContent = building;
                buildingsList.appendChild(li);
            });

            const resourcesList = section.querySelector('[data-field="resources-list"]');
            if (!spa.resources || spa.resources.length === 0) {

                const row = document.createElement("tr");

                const cell = document.createElement("td");
                cell.colSpan = 3;
                cell.textContent = "No resources are currently listed for this area.";

                row.appendChild(cell);
                resourcesList.appendChild(row);

                container.appendChild(section);
                return;
            }

            spa.resources.forEach(function (resource) {
                const row = document.createElement("tr");

                const resourceCell = document.createElement("td");
                resourceCell.textContent = resource.resource || "";

                const locationCell = document.createElement("td");
                locationCell.textContent = resource.location || "";

                const telephoneCell = document.createElement("td");
                telephoneCell.textContent = resource.telephone || "";

                row.appendChild(resourceCell);
                row.appendChild(locationCell);
                row.appendChild(telephoneCell);

                resourcesList.appendChild(row);
            });

            container.appendChild(section);
        });
    }

})();

