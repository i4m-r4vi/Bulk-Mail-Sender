const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");
const dotenv = require("dotenv");
const fs = require("fs");

dotenv.config();

const app = express();
const upload = multer({ dest: "uploads/" });

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static("public"));

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL,
    pass: process.env.PASS,
  },
});

app.get("/", (req, res) => {
  res.render("index");
});

app.post("/send-emails", upload.single("file"), async (req, res) => {
  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    let results = [];
    let sentEmails = [];

    for (const student of data) {
      const { Name, Email, Marks } = student;

      const mailOptions = {
        from: process.env.EMAIL,
        to: Email,
        subject: "Your Marks",
        text: `Hello ${Name},\n\nYour marks are: ${Marks}\n\nBest regards,\nTeam Techno Trivia`,
      };

      try {
        await transporter.sendMail(mailOptions);
        results.push({ name: Name, email: Email, status: "✅ Sent" });

        // Save only successful ones
        sentEmails.push({ Name, Email, Marks, Status: "Sent" });
      } catch (err) {
        results.push({ name: Name, email: Email, status: "❌ Failed" });
      }
    }

    // Export sent emails if any
    let downloadFile = null;
    if (sentEmails.length > 0) {
      const newWB = xlsx.utils.book_new();
      const newWS = xlsx.utils.json_to_sheet(sentEmails);
      xlsx.utils.book_append_sheet(newWB, newWS, "Sent Emails");

      const filename = `sent_emails_${Date.now()}.xlsx`;
      downloadFile = path.join("uploads", filename);
      xlsx.writeFile(newWB, downloadFile);
    }

    // Pass file name to result.ejs
    res.render("result", { results, downloadFile });
  } catch (error) {
    console.error(error);
    res.status(500).send("Something went wrong!");
  }
});

// Route for downloading exported file
app.get("/download/:file", (req, res) => {
  const filePath = path.join(__dirname, "uploads", req.params.file);
  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.status(404).send("File not found!");
  }
});

app.listen(3000, () => {
  console.log("🚀 Server running at http://localhost:3000");
});
