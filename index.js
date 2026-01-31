
const express = require("express");
const multer = require("multer");
const fs = require("fs-extra");
const path = require("path");
const cors = require("cors");
const { spawn } = require("child_process");
const { pipeline } = require("stream");
const dotenv = require("dotenv");

dotenv.config();

const PORT = process.env.PORT || 5000;
const SOFFICE_PATH =
  process.platform === "win32"
  ? "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
  : "soffice";

  


const app = express();
app.use(cors());


//  Prevent connection timeout 

app.use((req, res, next) => {
  req.setTimeout(0);
  res.setTimeout(0);
  next();
});


 // Directories
 
const uploadDir = path.join(__dirname, "uploads");
const outputDir = path.join(__dirname, "output");

fs.ensureDirSync(uploadDir);
fs.ensureDirSync(outputDir);


// Multer storage 

const storage = multer.diskStorage({
  destination: (_, __, cb) => cb(null, uploadDir),
  filename: (_, file, cb) => {
    const safeName =
      Date.now() + "_" + file.originalname.replace(/\s+/g, "_");
    cb(null, safeName);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
});


//  OFFICE → PDF (Word / PPT)
 
app.post(
  "/api/office-to-pdf",
  upload.single("file"),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      const inputPath = req.file.path;
      const pdfName = path.parse(req.file.path).name + ".pdf";
      const outputPath = path.join(outputDir, pdfName);

      await new Promise((resolve, reject) => {
  const process = spawn(SOFFICE_PATH, [
    "--headless",
    "--nologo",
    "--nofirststartwizard",
    "--convert-to",
    "pdf",
    "--outdir",
    outputDir,
    inputPath,
  ]);

  process.on("error", reject);
  process.on("close", (code) => {
    if (code === 0) resolve();
    else reject(new Error("LibreOffice failed"));
  });
});


      if (!(await fs.pathExists(outputPath))) {
        throw new Error("PDF not generated");
      }

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${pdfName}"`
      );

      pipeline(
  fs.createReadStream(outputPath),
  res,
  async (err) => {
    if (err) {
      console.error(" Stream failed", err);
    }

    await fs.remove(inputPath).catch(() => {});
    await fs.remove(outputPath).catch(() => {});
  }
);
    } catch (err) {
      console.error(" Conversion error:", err);
      res.status(500).json({ error: "Conversion failed" });
    }
  }
);

// MULTER ERROR HANDLER 
 
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    console.error(" Multer error:", err);
    return res.status(400).json({ error: err.message });
  }
  if (err) {
    console.error("Unknown error:", err);
    return res.status(500).json({ error: "Upload failed" });
  }
  next();
});

app.listen(PORT, () => {
  console.log(` Backend running at http://localhost:${PORT}`);
});


