// Server code
const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 3000;

// Middleware to parse JSON data from the form
app.use(bodyParser.json());
app.use(express.static('public')); // Serve the static files in the "public" folder

// Route to serve the HTML form
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Route to handle form submission and store data in Excel
app.post('/submit', (req, res) => {
    const formData = req.body;

    // Check if Excel file exists, otherwise create a new one
    const filePath = path.join(__dirname, 'student_data.xlsx');
    let workbook;
    let worksheet;

    if (fs.existsSync(filePath)) {
        // Read the existing file
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets['Students'];

        // If the 'Students' sheet doesn't exist, create it
        if (!worksheet) {
            worksheet = xlsx.utils.aoa_to_sheet([['Name', 'Date of Birth', 'Gender', 'Email', 'Phone', 'Address', 'City', 'State', 'Course', 'Hobbies', 'Department']]);
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Students');
        }
    } else {
        // Create a new workbook and worksheet
        workbook = xlsx.utils.book_new();
        worksheet = xlsx.utils.aoa_to_sheet([['Name', 'Date of Birth', 'Gender', 'Email', 'Phone', 'Address', 'City', 'State', 'Course', 'Hobbies', 'Department']]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Students');
    }

    // Ensure "Department" header is present
    const headers = ['Name', 'Date of Birth', 'Gender', 'Email', 'Phone', 'Address', 'City', 'State', 'Course', 'Hobbies', 'Department'];
    const headerRow = xlsx.utils.sheet_to_json(worksheet, { header: 1 })[0];
    if (!headerRow || headerRow.length < headers.length) {
        xlsx.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A1' });
    }

    // Prepare data row including "Department"
    const dataRow = [
        formData.name,
        formData.dob,
        formData.gender,
        formData.email,
        formData.phone,
        formData.address,
        formData.city,
        formData.state,
        formData.course,
        formData.hobbies,
        formData.department  // Include department here
    ];

    // Determine the next row number
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    const nextRow = range.e.r + 2; // e.r gives the last row index, so +2 for the next row

    // Append the new data row at the next available row
    xlsx.utils.sheet_add_aoa(worksheet, [dataRow], { origin: `A${nextRow}` });

    // Update the reference range to include the new data
    worksheet['!ref'] = xlsx.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: nextRow - 1, c: headers.length - 1 }
    });

    // Write the updated workbook to the file
    xlsx.writeFile(workbook, filePath);

    // Send response to the frontend
    res.json({ message: 'Form submitted successfully!' });
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

