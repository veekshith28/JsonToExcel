const fs = require("fs");
const XLSX = require("xlsx");

let sectionCounter = 0;

// Function to flatten nested objects
function iterateObject(obj, prefix = "", parentKey = "") {
  let flattened = {};

  for (let key in obj) {
    const fullPath = parentKey ? `${parentKey}.${key}` : key;

    if (Array.isArray(obj[key])) {
      flattened[`${prefix}${key}`] = obj[key].join(", "); // Flatten arrays
    } else if (typeof obj[key] === "object" && obj[key] !== null) {
      const nestedObj = iterateObject(obj[key], `${prefix}${key}.`, fullPath);
      flattened = { ...flattened, ...nestedObj };
      flattened[`${prefix}${key}`] = `SHEET::${fullPath}`;
    } else {
      if (key === "sections" || key === "test") {
        flattened[`${prefix}${key}`] = `SHEET::${key}`;
      } else {
        flattened[`${prefix}${key}`] = obj[key];
      }
    }
  }

  return flattened;
}
// Function to process the nested data and convert to Excel format
function processJSONToExcel(data) {
  const workbook = XLSX.utils.book_new();
  let sectionIndex = -1;

  try {
    const mainData = { ...data };
    delete mainData.sections;
    delete mainData.test;

    const mainWorksheet = XLSX.utils.json_to_sheet([iterateObject(mainData)]);
    XLSX.utils.book_append_sheet(workbook, mainWorksheet, "Sheet1");

    // Handling 'test' data
    if (data.test && typeof data.test === "object") {
      const testWorksheet = XLSX.utils.json_to_sheet([
        iterateObject(data.test),
      ]);
      XLSX.utils.book_append_sheet(workbook, testWorksheet, "test_obj");

      // Handling nested 'test' data
      for (let key in data.test) {
        if (typeof data.test[key] === "object" && data.test[key] !== null) {
          const nestedTestWorksheet = XLSX.utils.json_to_sheet([
            iterateObject(data.test[key]),
          ]);
          XLSX.utils.book_append_sheet( workbook,nestedTestWorksheet, `test_obj.${key}`
          );
        }
      }
    }

    // Handling 'sections' data
    if (data.sections && Array.isArray(data.sections)) {
      data.sections.forEach((section) => {
        sectionIndex++;

        let booksData = [];
        if (section.books && Array.isArray(section.books)) {
          section.books.forEach((book) => {
            sectionCounter++;
            booksData.push(iterateObject(book));
          });

          const booksWorksheet = XLSX.utils.json_to_sheet(booksData);
          XLSX.utils.book_append_sheet(workbook,booksWorksheet,`sections_arr.${sectionIndex}.books_arr`);
        }
      });
    }

    const outputFilePath = "output.xlsx";
    XLSX.writeFile(workbook, outputFilePath, { bookType: "xlsx" });
    console.log(`Excel file "${outputFilePath}" has been created.`);
  } catch (error) {
    console.error("Error processing JSON data:", error);
  }
}

// Read the JSON file
fs.readFile("input.json", "utf8", (err, jsonString) => {
  if (err) {
    console.error("Error reading JSON file:", err);
    return;
  }

  try {
    const jsonData = JSON.parse(jsonString);
    processJSONToExcel(jsonData);
  } catch (error) {
    console.error("Error parsing JSON data:", error);
  }
});
