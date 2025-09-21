const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const upload = multer({ dest: "uploads/" });

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static("public"));


const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: `${process.env.EMAIL}`,     
    pass: `${process.env.PASS}`,        
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

    for (const student of data) {
      const { Name, Email, Marks } = student;

      const mailOptions = {
        from: "your_email@gmail.com",
        to: Email,
        subject: "Your Marks",
        text: `Hello ${Name},\n\nYour marks are: ${Marks}\n\nBest regards,\nTeam\nTechno Trivia`,
      };

      try {
        await transporter.sendMail(mailOptions);
        results.push({ name: Name, email: Email, status: "✅ Sent" });
      } catch (err) {
        results.push({ name: Name, email: Email, status: "❌ Failed" });
      }
    }

    res.render("result", { results });
  } catch (error) {
    console.error(error);
    res.status(500).send("Something went wrong!");
  }
});

app.listen(3000, () => {
  console.log("🚀 Server running at http://localhost:3000");
});
