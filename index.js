const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path");
const ImageModule = require("open-docxtemplater-image-module");

// Load the docx file as binary content
const content = fs.readFileSync(path.resolve(__dirname, "aayamSlip.docx"), "binary");

// Unzip the content of the file
const zip = new PizZip(content);

// Configure the image module
const imageModule = new ImageModule({
    centered: true, // Set to true if you want to center the image
    fileType: "docx", // The file type being processed
    getImage: function (tagValue) {
        return fs.readFileSync(tagValue); // Read image file as a buffer
    },
    getSize: function (imgBuffer, tagValue) {
        return [200, 150]; // Width and height in pixels
    },
});

// Create the document with the image module
const doc = new Docxtemplater(zip, {
    modules: [imageModule], // Add image module
    paragraphLoop: true,
    linebreaks: true,
});

// Replace placeholders with data, including the image path
doc.render({
    name: "sujal",
    evenName: "Tech Expo",
    amount: "300",
    payMethod: "Cash",
    image: path.resolve(__dirname, "image.png"), // ✅ Make sure this path is correct
});

// Generate the document as a buffer
const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

// Write the file
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);

console.log("✅ Document generated successfully with an image!");
