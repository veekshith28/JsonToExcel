const fs = require("fs"); // File system module for reading files
const XLSX = require("xlsx"); // XLSX module for Excel file manipulation

// Function to flatten nested objects
function iterateObject(obj, parentPath = "") {
  let flattened = {};

  // Iterate through the object's keys
  for (let key in obj) {
    if (Array.isArray(obj[key])) {
      flattened[`${parentPath}${key}`] = obj[key].join(", "); // Flatten arrays
    } else if (typeof obj[key] === "object" && obj[key] !== null) {
      // Recursively call the function for nested objects
      const nestedObj = iterateObject(obj[key], `${parentPath}${key}.`);
      flattened = { ...flattened, ...nestedObj };
    } else {
      // For non-object values, assign them directly
      flattened[`${parentPath}${key}`] = obj[key];
    }
  }

  return flattened; // Return the flattened object
}

// Function to process the JSON data and convert it into an Excel file
function processJSONToExcel(data) {
  const workbook = XLSX.utils.book_new(); // Create a new Excel workbook

  try {
    let uniqueKeys = new Set(); // Initialize a set to store unique keys

    // Function to extract unique keys from nested objects
    const extractKeys = (obj) => {
      for (let key in obj) {
        if (
          typeof obj[key] === "object" &&
          obj[key] !== null &&
          !Array.isArray(obj[key])
        ) {
          uniqueKeys.add(key); // Add unique keys to the set
          extractKeys(obj[key]); // Recursively extract keys from nested objects
        }
      }
    };

    extractKeys(data); // Call the key extraction function for the input JSON data

    const headers = Array.from(uniqueKeys); // Convert the set of unique keys to an array

    // Function to process nested objects and create Excel sheets
    const processObjects = (obj, parentPath = "") => {
      for (let key in obj) {
        if (
          typeof obj[key] === "object" &&
          obj[key] !== null &&
          !Array.isArray(obj[key])
        ) {
          // Generate sheet name based on the object hierarchy
          const sheetName = parentPath ? `${parentPath}_${key}` : key;
          const sheetData = [iterateObject(obj[key])]; // Flatten nested object data
          // Create an Excel sheet with flattened data and specified headers
          const nestedWorksheet = XLSX.utils.json_to_sheet(sheetData, {
            header: headers,
          });
          // Append the sheet to the workbook
          XLSX.utils.book_append_sheet(workbook, nestedWorksheet, sheetName);
          processObjects(obj[key], sheetName); // Recursively process nested objects
        }
      }
    };

    processObjects(data); // Start processing nested objects from the root

    const outputFilePath = "output.xlsx"; // Define the output file path
    XLSX.writeFile(workbook, outputFilePath, { bookType: "xlsx" }); // Write the workbook to an Excel file
    console.log(`Excel file "${outputFilePath}" has been created.`); // Log success message
  } catch (error) {
    console.error("Error processing JSON data:", error); // Handle processing errors
  }
}

// Read the JSON file
fs.readFile("input.json", "utf8", (err, jsonString) => {
  if (err) {
    console.error("Error reading JSON file:", err); // Handle file reading errors
    return;
  }

  try {
    const jsonData = JSON.parse(jsonString); // Parse the JSON data
    processJSONToExcel(jsonData); // Process the parsed JSON data to generate an Excel file
  } catch (error) {
    console.error("Error parsing JSON data:", error); // Handle JSON parsing errors
  }
});
