const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-pwndoc'); // Use the forked version
const sizeOf = require('image-size');

const app = express();
const port = 3000;

// Configure multer for file uploads (using memory storage)
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 }, // Limit file size to 50MB (adjust as needed)
});

// Endpoint to modify the Word document
app.post('/modifyReport', upload.fields([
    { name: 'wordFile', maxCount: 1 },
    { name: 'image', maxCount: 10 }
]), async (req, res) => {
    try {
        console.log('Request received...');
        const wordFile = req.files['wordFile']?.[0];
        if (!wordFile) {
            console.error('Word file is not present.');
            return res.status(400).send('No Word document provided.');
        }

        console.log('Word file details:', wordFile);

        const content = wordFile.buffer;
        const zip = new PizZip(content);

        // Set up the image module
        const imageModule = new ImageModule({
            centered: false,
            fileType: 'docx',
            getImage: (tagValue) => {
                console.log(`Looking for image with tag: ${tagValue}`);
                // Find the image file in the uploaded files based on tagValue
                const image = req.files['image']?.find(img => img.originalname === '1.jpg');
                if (!image) {
                    throw new Error(`Image file ${tagValue} not found`);
                }
                console.log(`Image found: ${image.originalname}`);
                return image.buffer; // Return the image buffer directly
            },
            getSize: (imgBuffer) => {
                // Get image dimensions using image-size
                const size = sizeOf(imgBuffer);
                console.log(`Image size: ${size.width}x${size.height}`);
                return [size.width, size.height];
            }
        });

        // Initialize docxtemplater with the zip content and attach the image module
        const doc = new Docxtemplater()
            .attachModule(imageModule)
            .loadZip(zip);

        // Assuming a single placeholder called {%image1} in your DOCX
        const imagePlaceholders = { image1: '1.jpg' };
        console.log('Image placeholders:', imagePlaceholders);

        doc.setData(imagePlaceholders);

        // Render the document with the new images
        try {
            doc.render();
        } catch (renderError) {
            console.error('Error during document rendering:', renderError);
            return res.status(500).send('Error rendering the document');
        }

        // Generate the modified document buffer
        const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });

        // Set headers and send the modified document as the response
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=modified_report.docx');
        res.send(outputBuffer);

        console.log('Response sent successfully.');
    } catch (error) {
        console.error('Error processing the request:', error);
        res.status(500).send('Error processing the document');
    }
});

// Start the server
app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});
