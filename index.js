const fs = require("fs"); // File system module for reading files
const XLSX = require("xlsx"); // XLSX module for Excel file manipulation

let sectionCounter = 0;

// Function to iterate nested objects
function iterateObject(obj, prePath = "", parentPath = "") {
  let flattened = {};

  // Iterate through the object's keys
  for (let key in obj) {
    const fullPath = parentPath ? `${parentPath}.${key}` : key; // Create full path for nested objects

    // Check if the value is an array
    if (Array.isArray(obj[key])) {
      flattened[`${prePath}${key}`] = obj[key].join(", "); // Flatten arrays
    } else if (typeof obj[key] === "object" && obj[key] !== null) {
      // Recursively call the function for nested objects
      const nestedObj = iterateObject(obj[key], `${prePath}${key}.`, fullPath);
      flattened = { ...flattened, ...nestedObj };
      flattened[`${prePath}${key}`] = `SHEET::${fullPath}`; // Indicate nested objects
    } else {
      // Check specific keys and assign values accordingly
      if (key === "sections" || key === "test") {
        flattened[`${prePath}${key}`] = `SHEET::${key}`; // Indicate sections or test objects
      } else {
        flattened[`${prePath}${key}`] = obj[key]; // Assign regular values
      }
    }
  }

  return flattened; // Return the object
}

// Function to process the nested data and convert to Excel format
function JsonToCSV(data) {
  const workbook = XLSX.utils.book_new(); // Create a new workbook
  let sectionIndex = -1; // Initialize section index

  try {
    // Create a copy of the data object excluding 'sections' and 'test'
    const mainData = { ...data };
    delete mainData.sections;
    delete mainData.test;

    // Convert the main data to a worksheet and append it to the workbook
    const mainWorksheet = XLSX.utils.json_to_sheet([iterateObject(mainData)]);
    XLSX.utils.book_append_sheet(workbook, mainWorksheet, "Sheet1");

    // Handling 'test' data
    if (data.test && typeof data.test === "object") {
      // Convert 'test' data to a worksheet and append it to the workbook
      const testWorksheet = XLSX.utils.json_to_sheet([
        iterateObject(data.test),
      ]);
      XLSX.utils.book_append_sheet(workbook, testWorksheet, "test_obj");

      // Handling nested 'test' data
      for (let key in data.test) {
        if (typeof data.test[key] === "object" && data.test[key] !== null) {
          // Convert nested 'test' data to a worksheet and append it to the workbook
          const nestedTestWorksheet = XLSX.utils.json_to_sheet([
            iterateObject(data.test[key]),
          ]);
          XLSX.utils.book_append_sheet(
            workbook,
            nestedTestWorksheet,
            `test_obj.${key}`
          );
        }
      }
    }

    // Handling 'sections' data
    if (data.sections && Array.isArray(data.sections)) {
      // Iterate through 'sections' array
      data.sections.forEach((section) => {
        sectionIndex++; // Increment section index

        let booksData = [];
        // Check for 'books' array in each section
        if (section.books && Array.isArray(section.books)) {
          // Iterate through 'books' array
          section.books.forEach((book) => {
            sectionCounter++; // Increment section counter
            booksData.push(iterateObject(book)); // Push flattened book data
          });

          // Convert 'books' data to a worksheet and append it to the workbook
          const booksWorksheet = XLSX.utils.json_to_sheet(booksData);
          XLSX.utils.book_append_sheet(
            workbook,
            booksWorksheet,
            `sections_arr.${sectionIndex}.books_arr`
          );
        }
      });
    }

    // Define the output file path
    const outputFilePath = "output.xlsx";
    // Write the workbook to the output file
    XLSX.writeFile(workbook, outputFilePath, { bookType: "xlsx" });
    console.log(`Excel file "${outputFilePath}" has been created.`);
  } catch (error) {
    // Handle errors
    console.error("Error processing JSON data:", error);
  }
}

// Read the JSON file
fs.readFile("input.json", "utf8", (err, jsonString) => {
  if (err) {
    // Handle file reading errors
    console.error("Error reading JSON file:", err);
    return;
  }

  try {
    const jsonData = JSON.parse(jsonString); // Parse JSON data
    JsonToCSV(jsonData); // Call function to process JSON data
  } catch (error) {
    // Handle JSON parsing errors
    console.error("Error parsing JSON data:", error);
  }
});
