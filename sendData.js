const xlsx = require('xlsx');
const axios = require('axios');

// Function to validate email format
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Function to read Excel file and return data as array of objects
function readExcel(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet);
}

// Function to make API request with provided data
async function sendDataToAPI(data) {
    const apiUrl = 'https://www.techshotsapp.com/api/cms/newsletterSubscriber';
    let count = 0;
    for (const row of data) {
        // Check if name and email are not null, undefined, or empty strings
        if (row.name && row.email && isValidEmail(row.email)) {
            const payload = {
                name: row.name.trim(),
                email: row.email.trim()
            };
            count++;
            console.log(payload);
            try {
                await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for 3 seconds
                const response = await axios.post(apiUrl, payload);

                if (response.status === 200) {
                    console.log(`Data sent successfully for ${row.name}`);
                } else {
                    console.error(`Failed to send data for ${row.name}: ${response.statusText}`);
                }
            } catch (error) {
                console.error(`Error sending data for ${row.name}:`, error.message);
            }
        } else {
            console.log('Skipping row due to missing name, email, or invalid email format:', row);
        }
    }
    console.log(`Total number of valid rows: ${count} and data length is: ${data.length}`);
}

// Main function to read Excel file and send data to API
function main() {
    const filePath = 'datashort.xlsx'; // Path to your Excel file
    const data = readExcel(filePath);
    sendDataToAPI(data);
}

main();