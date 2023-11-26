// server.js

const express = require("express");
const nodemailer = require("nodemailer");
const exceljs = require("exceljs");
require("dotenv").config(); // Load environment variables

const cors = require("cors");

const app = express();
app.use(cors());

const port = process.env.PORT || 3001;

// Middleware to parse JSON data
app.use(express.json());

// Nodemailer transporter setup
const transporter = nodemailer.createTransport({
	// host: "smtp.ethereal.email",
	// port: 587,
	service: "gmail",
	auth: {
		user: process.env.EMAIL_USER, // Your Gmail address
		pass: process.env.EMAIL_PASSWORD, // App Password generated in the previous step
	},
	// auth: {
	// 	user: "oleta.jast@ethereal.email",
	// 	pass: "VY4dHzhY7UDRAeExf5",
	// },
});

// API endpoint to receive survey results
app.post("/api/send-survey-results", async (req, res) => {
	const { allAnswers, correctAnswers } = req.body;

	// Generate Excel files
	const surveyResultsWorkbook = createExcelWorkbook(allAnswers, correctAnswers);

	// Send emails with attachments
	await sendEmail("dahalprastut@gmail.com", "Survey Results", surveyResultsWorkbook);

	res.status(200).json({ success: true });
});

// Helper function to create Excel workbook
function createExcelWorkbook(allAnswers, correctAnswers) {
	const workbook = new exceljs.Workbook();

	// Create worksheet for All Answers
	const allAnswersWorksheet = workbook.addWorksheet("All Answers");
	createWorksheetFromData(allAnswersWorksheet, allAnswers);

	// Create worksheet for Correct Answers
	const correctAnswersWorksheet = workbook.addWorksheet("Correct Answers");
	createWorksheetFromData(correctAnswersWorksheet, correctAnswers);

	return workbook;
}

function createWorksheetFromData(worksheet, data) {
	// Adding headers
	const headers = Object.keys(data[0]);
	worksheet.addRow(headers);

	// Adding data rows
	data.forEach((row) => {
		const rowData = headers.map((header) => row[header]);
		worksheet.addRow(rowData);
	});
}

// Helper function to send email with attachment
async function sendEmail(to, subject, attachment) {
	const mailOptions = {
		// from: "oleta.jast@ethereal.email", // Update with your email
		from: "prastutdahal2717@gmail.com", // Update with your email
		to,
		subject,
		text: "Survey Results",
		attachments: [
			{
				filename: `${subject}.xlsx`,
				content: await attachment.xlsx.writeBuffer(),
				encoding: "base64",
			},
		],
	};

	return transporter.sendMail(mailOptions);
}

app.listen(port, () => {
	console.log(`Server is running on port ${port}`);
});
