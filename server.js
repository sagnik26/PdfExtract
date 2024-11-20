import express from "express";
import multer from "multer";
import { exec } from "child_process";
import { join } from "path";
import { unlinkSync } from "fs";
import { fileURLToPath } from "url";
import path from "path";

// Define __dirname for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ dest: "uploads/" });

// Serve static HTML file
// app.use(express.static("public"));
// Serve static HTML file
app.use(express.static(path.join(__dirname, "public")));

// Endpoint to handle file upload and processing
app.post("/convert", upload.single("pdf"), (req, res) => {
  const pdfPath = req.file.path;
  const outputPath = join(__dirname, "output", `${req.file.filename}.xlsx`);

  const pythonScript = join(__dirname, "convert_pdf_to_excel.py");
  const command = `python3 ${pythonScript} "${pdfPath}" "${outputPath}"`;

  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error(`Error: ${stderr}`);
      res.status(500).send("Error processing the PDF.");
      return;
    }

    console.log("STD OUT", stdout);

    // Send the Excel file to the client
    res.download(outputPath, `converted.xlsx`, (err) => {
      if (err) console.error(err);

      // Clean up files
      unlinkSync(pdfPath);
      unlinkSync(outputPath);
    });
  });
});

// Start the server
app.listen(3000, () => {
  console.log("Server is running on http://localhost:3000");
});
