const fs = require('fs');
const path = require('path');
const unzipper = require('unzipper');
const xlsx = require('xlsx');

// Paths
const zipFilePath = 'zip_folder/images.zip';  // Path to your zip file
const outputDirectory = 'zip_folder/extracted_files';  // Ensure this is a subdirectory to avoid reading the zip file itself
const excelFilePath = 'zip_folder/images_name.xlsx';  // Output path for the Excel file

// Function to extract zip file
async function extractZip(filePath, outputDir) {
    await fs.createReadStream(filePath)
        .pipe(unzipper.Extract({ path: outputDir }))
        .promise();
}

// Function to get filenames from a directory recursively
function getFilenames(directory) {
    let filenames = [];
    function readDirRecursively(dir) {
        const files = fs.readdirSync(dir);
        files.forEach(file => {
            const fullPath = path.join(dir, file);
            if (fs.lstatSync(fullPath).isDirectory()) {
                readDirRecursively(fullPath);
            } else {
                filenames.push(fullPath);
            }
        });
    }
    readDirRecursively(directory);
    return filenames;
}

// Function to process filenames and extract product names and extensions
function extractProductNamesAndExtensions(filenames) {
    return filenames.map(filename => {
        const parsedPath = path.parse(filename);
        return {
            name: parsedPath.name,
            extension: parsedPath.ext
        };
    });
}

// Function to save product names and extensions to an Excel file
function saveToExcel(productData, filePath) {
    const worksheetData = [['Product Name', 'Extension'], ...productData.map(item => [item.name, item.extension])];
    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Products');
    xlsx.writeFile(workbook, filePath);
}

// Main function
async function main() {
    try {
        // Ensure the output directory exists
        if (!fs.existsSync(outputDirectory)) {
            fs.mkdirSync(outputDirectory, { recursive: true });
        }

        // Extract the zip file
        await extractZip(zipFilePath, outputDirectory);
        console.log('Zip file extracted successfully.');

        // Get filenames from the output directory
        const filenames = getFilenames(outputDirectory);

        // Extract product names and extensions from filenames
        const productData = extractProductNamesAndExtensions(filenames);

        // Save product names and extensions to Excel
        saveToExcel(productData, excelFilePath);

        console.log(`Product names and extensions have been extracted and saved to ${excelFilePath}`);
    } catch (error) {
        console.error('Error:', error);
    }
}

// Run the main function
main();
