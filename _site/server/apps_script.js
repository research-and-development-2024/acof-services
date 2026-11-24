function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ACOF-Services Website Tools")
    .addItem("Publish Resources", "updateResources")
    .addToUi();
}

function updateResources() {
  commitResourcesToGitHub();
}

function getGitHubFileMetadata_(apiUrl, token, branch) {
  const url = apiUrl + "?ref=" + encodeURIComponent(branch);

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status === 404) {
    return null;
  }

  if (status !== 200) {
    throw new Error(
      "Could not check existing GitHub file. Status: " +
      status +
      "\n" +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}

function commitResourcesToGitHub() {
  const props = PropertiesService.getScriptProperties();

  const REPO_OWNER = props.getProperty("REPO_OWNER");
  const REPO = props.getProperty("REPO") || "acof-services";
  const REPO_TOKEN = props.getProperty("REPO_TOKEN");
  const REPO_BRANCH = props.getProperty("REPO_BRANCH") || "main";

  if (!REPO_OWNER || !REPO || !REPO_TOKEN) {
    throw new Error("Missing REPO_OWNER or REPO_TOKEN in Script Properties.");
  }

  const filePath = "data/resources.json";

  const data = buildSPAData();
  
  // Converts a data object into a neatly indented, 2-space formatted JSON string for saving.
  const fileContent = JSON.stringify(data, null, 2);

  const apiUrl =
    "https://api.github.com/repos/" +
    encodeURIComponent(REPO_OWNER) + "/" +
    encodeURIComponent(REPO) +
    "/contents/" +
    filePath.split("/").map(encodeURIComponent).join("/");

  const existingFile = getGitHubFileMetadata_(apiUrl, REPO_TOKEN, REPO_BRANCH);

  const payload = {
    message: "Updating SPA resources in JSON from Apps Script",
    content: Utilities.base64Encode(fileContent),
    branch: REPO_BRANCH
  };

  if (existingFile && existingFile.sha) {
    payload.sha = existingFile.sha;
  }

  const response = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + REPO_TOKEN,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status !== 200 && status !== 201) {
    throw new Error("GitHub commit failed. Status: " + status + "\n" + body);
  }

  SpreadsheetApp.getUi().alert("Tenant resources successfully updated!");
}

/**
 * Builds a JSON object containing all SPA information.
 *
 * Assumptions:
 * 1. There is a sheet named "SPA Directory".
 * 2. The SPA Directory sheet contains the columns:
 *      spa-id
 *      spa-display-name
 *      buildings
 *
 * 3. For each spa-id, there is a sheet with exactly the same name.
 *    Example:
 *      spa-1
 *      spa-2
 *      spa-3
 *
 * 4. Each SPA sheet contains the columns:
 *      Resource
 *      Location
 *      Telephone
 *      Notes
 *
 * Returns:
 * [
 *   {
 *     "spa-id": "spa-2",
 *     "display-name": "SPA 2 – San Fernando Valley",
 *     "buildings": ["Casa Del Sol", "Cornerstone"],
 *     "resources": [
 *       {
 *         "resource": "...",
 *         "location": "...",
 *         "telephone": "..."
 *       }
 *     ]
 *   }
 * ]
 */
function buildSPAData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("buildSPAData: Reading SPA directory.");

  const directorySheet = ss.getSheetByName("SPA Directory");
  const directoryData = directorySheet.getDataRange().getValues();

  const headers = directoryData[0];

  const spaIdCol = headers.indexOf("spa-id");
  const displayNameCol = headers.indexOf("spa-display-name");
  const buildingsCol = headers.indexOf("buildings");

  const result = [];

  for (let i = 1; i < directoryData.length; i++) {
    const row = directoryData[i];

    const spaId = row[spaIdCol];
    const displayName = row[displayNameCol];
    const buildingsText = row[buildingsCol] || "";

    const buildings = buildingsText
      .toString()
      .split("\n")
      .map(b => b.trim())
      .filter(b => b.length > 0);

    const spaSheet = ss.getSheetByName(spaId);
    Logger.log("buildSPAData: Reading resources for " + spaId);

    if (!spaSheet) {
      continue;
    }

    const resourcesData = spaSheet.getDataRange().getValues();

  // If the SPA sheet contains only the header row (no resources),
  // include the SPA in the JSON output with an empty resources array
    if (resourcesData.length < 2) {
      result.push({
        "spa-id": spaId,
        "display-name": displayName,
        "buildings": buildings,
        "resources": []
      });

      continue;
    }

    const resourceHeaders = resourcesData[0];

    const resourceCol = resourceHeaders.indexOf("Resource");
    const locationCol = resourceHeaders.indexOf("Location");
    const telephoneCol = resourceHeaders.indexOf("Telephone");

    const resources = [];
    for (let j = 1; j < resourcesData.length; j++) {

      const resourceRow = resourcesData[j];
      const resourceName = resourceRow[resourceCol];

      // Ignore empty rows
      if (!resourceName) {
        continue;
      }

      resources.push({
        "resource": resourceRow[resourceCol],
        "location": resourceRow[locationCol],
        "telephone": resourceRow[telephoneCol]
      });
    }

    result.push({
      "spa-id": spaId,
      "display-name": displayName,
      "buildings": buildings,
      "resources": resources
    });
  }

  return result;
}

function testBuildSPA() {
  Logger.log("testBuildSPA: Initiating testing.");
  const data = buildSPAData();
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Debug");

  sheet.getRange("A1").setValue(
    JSON.stringify(data, null, 2)
  );

  Logger.log("testBuildSPA: Testing completed.");
}