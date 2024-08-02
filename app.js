const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

// Set EJS as the templating engine
app.set('view engine', 'ejs');

// Use body-parser middleware
app.use(bodyParser.urlencoded({ extended: true }));

// Serve static files from the 'temp' directory
app.use('/temp', express.static(path.join(__dirname, 'temp')));

// Function to fill Word template with form data
async function fillTemplateWithData(templatePath, data) {
    try {
        // Read the original Word document
        const originalDocx = fs.readFileSync(templatePath, 'binary');

        // Open the original document for further modifications
        const zip = new PizZip(originalDocx);
        const doc = new Docxtemplater().loadZip(zip);

        // Set the data to replace placeholders
        doc.setData(data);

        // Apply the data to the document template
        doc.render();

        // Return the modified content
        return doc.getZip().generate({ type: 'nodebuffer' });
    } catch (error) {
        console.error('Error filling template with data:', error);
        throw error;
    }
}

// Define a route for the home page
app.get('/', (req, res) => {
    res.render('index'); // Render the index.ejs file
});

// Define a route for form submission
app.post('/generate', async (req, res) => {
    console.log('Form submission received...');
    const { firstName, lastName } = req.body;

    try {
        const templatePath = path.join(__dirname, 'templates', 'template.docx');

        // Get current date
        const currentDate = new Date();

        // Format the date
        const formattedDay = getFormattedDay(currentDate.getDate());
        const formattedMonth = getCurrentMonthName(currentDate.getMonth());

        // Fill the template with data and generate the Word document
        const generatedDocx = await fillTemplateWithData(templatePath, { firstName, lastName, date: formattedDay, month: formattedMonth });

        // Save the generated document to a temporary file
        const tempFileName = `NDA_${firstName}_${lastName}.docx`;
        const tempFilePath = path.join(__dirname, 'temp', tempFileName);
        fs.writeFileSync(tempFilePath, generatedDocx);

        console.log(`File created: ${tempFilePath}`);

        // Render the success page with the download link
        res.render('success', { tempFileName });
    } catch (error) {
        console.error('Error processing form submission:', error);
        res.status(500).send('Error generating document');
    }
});

// Route to directly download the file
app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filepath = path.join(__dirname, 'temp', filename);

    if (fs.existsSync(filepath)) {
        res.download(filepath);
    } else {
        res.status(404).send('File not found');
    }
});

// Function to get the formatted day (e.g., 1st, 2nd, 3rd)
function getFormattedDay(day) {
    if (day >= 11 && day <= 13) {
        return day + 'th';
    }
    switch (day % 10) {
        case 1: return day + 'st';
        case 2: return day + 'nd';
        case 3: return day + 'rd';
        default: return day + 'th';
    }
}

// Function to get the name of the current month
function getCurrentMonthName(monthIndex) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months[monthIndex];
}

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server started on port ${PORT}`);
});
