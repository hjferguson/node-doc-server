// server.js
const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const Jimp = require('jimp');

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
        console.log('Files received:', req.files);

        // Check if the Word document is provided
        const wordFile = req.files['wordFile']?.[0];
        if (!wordFile) {
            console.error('Word file is not present.');
            return res.status(400).send('No Word document provided.');
        }

        console.log('Word file details:', wordFile);

        // Use the buffer directly from the uploaded file
        const content = wordFile.buffer;
        const zip = new PizZip(content);

        // Initialize docxtemplater with the zip content
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        console.log('Document loaded successfully.');

        // Process the images provided in the request
        const images = req.files['image'] || [];
        const processedImages = await Promise.all(images.map(async (img) => {
            try {
                if (!img.buffer) {
                    console.error(`Image buffer for ${img.originalname} is undefined.`);
                    return null;
                }

                // Use Jimp to read and resize the image buffer
                const image = await Jimp.read(img.buffer);
                const resizedBuffer = await image.resize(200, 200).getBufferAsync(Jimp.MIME_PNG);
                return resizedBuffer;
            } catch (error) {
                console.error(`Error processing image ${img.originalname}:`, error);
                return null;
            }
        }));

        console.log('All images processed.');

        // Assuming you want to add each image at the end of the document
        processedImages.forEach((imgBuffer, index) => {
            if (imgBuffer) {
                // This is where you could potentially add the image to the document
                // You'll need to modify this based on how you want to add images to docxtemplater
                // Currently, the code sets data placeholders, assuming further handling elsewhere
                doc.setData({
                    [`image${index + 1}`]: imgBuffer.toString('base64'),
                });
            }
        });

        // Generate the modified document buffer
        doc.render();
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
