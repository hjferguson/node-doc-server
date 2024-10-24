const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-pwndoc');
const sizeOf = require('image-size');

const app = express();
const port = 3000;

// Configure multer for file uploads (using memory storage)
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 }, // Limit file size to 50MB (adjust as needed)
});

// Function to add multiple placeholders directly into the DOCX XML based on number of images
function addPlaceholdersToXML(zip, imageCount) {
    const documentXml = zip.file('word/document.xml').asText();
    let placeholders = '';

    // Add a placeholder paragraph for each image
    for (let i = 1; i <= imageCount; i++) {
        placeholders += `
            <w:p>
                <w:r>
                    <w:t>{%image${i}}</w:t>
                </w:r>
            </w:p>
        `;
    }

    // Insert the placeholders at the end of the document body
    const modifiedXml = documentXml.replace('</w:body>', `${placeholders}</w:body>`);
    zip.file('word/document.xml', modifiedXml);
}

// Endpoint to modify the Word document
app.post('/modifyReport', upload.fields([
    { name: 'wordFile', maxCount: 1 },
    { name: 'image', maxCount: 15 } // Adjust maxCount as needed
]), async (req, res) => {
    try {
        console.log('Request received...');
        const wordFile = req.files['wordFile']?.[0];
        if (!wordFile) {
            console.error('Word file is not present.');
            return res.status(400).send('No Word document provided.');
        }

        console.log('Word file details:', wordFile);

        const images = req.files['image'] || [];
        console.log(`Number of images received: ${images.length}`);

        const content = wordFile.buffer;
        const zip = new PizZip(content);

        // Add the placeholders based on the number of images
        addPlaceholdersToXML(zip, images.length);

        // Set up the image module
        const imageModule = new ImageModule({
            centered: false,
            fileType: 'docx',
            getImage: (tagValue) => {
                console.log(`Looking for image with tag: ${tagValue}`);
                // Extract the image index from the tag (e.g., image1, image2, etc.)
                const imageIndex = parseInt(tagValue.replace('image', ''), 10) - 1;
                const image = images[imageIndex];
                if (!image) {
                    console.error(`Image for ${tagValue} not found.`);
                    throw new Error(`Image for ${tagValue} not found.`);
                }
                console.log(`Image found: ${image.originalname}`);
                return image.buffer;
            },
            getSize: (imgBuffer) => {
                const size = sizeOf(imgBuffer);
                const maxWidth = 400; // Set the maximum width (in pixels) for the image

                // Calculate new dimensions to maintain aspect ratio if width exceeds maxWidth
                if (size.width > maxWidth) {
                    const aspectRatio = size.height / size.width;
                    const newHeight = Math.round(maxWidth * aspectRatio);
                    console.log(`Resizing image to: ${maxWidth}x${newHeight}`);
                    return [maxWidth, newHeight];
                }

                console.log(`Original image size retained: ${size.width}x${size.height}`);
                return [size.width, size.height];
            }
        });

        // Initialize docxtemplater with the zip content and attach the image module
        const doc = new Docxtemplater()
            .attachModule(imageModule)
            .loadZip(zip);

        // Create dynamic data for images
        const imagePlaceholders = {};
        images.forEach((img, index) => {
            const placeholderName = `image${index + 1}`;
            imagePlaceholders[placeholderName] = img.originalname;
        });

        console.log('Image placeholders:', imagePlaceholders);

        // Set data with image placeholders
        doc.setData(imagePlaceholders);

        try {
            console.log('Rendering the document...');
            doc.render();
            console.log('Document rendered successfully.');
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
