const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;
const FILE_PATH = path.join(__dirname, 'submissions.xlsx');

app.use(cors());
app.use(express.json());

// Helper function to append data to Excel
async function appendDataToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  let worksheet;

  if (fs.existsSync(FILE_PATH)) {
    await workbook.xlsx.readFile(FILE_PATH);
    worksheet = workbook.getWorksheet('Submissions');
    if (!worksheet) {
      worksheet = workbook.addWorksheet('Submissions');
    }
  } else {
    worksheet = workbook.addWorksheet('Submissions');
  }

  // Add headers if worksheet is empty
  if (worksheet.rowCount === 0) {
    worksheet.addRow(['Name', 'College', 'Dream Job', 'Expected Salary', 'Experience', 'Projects']);
  }

  // Append the new row
  worksheet.addRow([
    data.name,
    data.college,
    data.dreamJob,
    data.expectedSalary,
    data.experience,
    data.projects,
  ]);

  await workbook.xlsx.writeFile(FILE_PATH);
}

// POST route to receive form data
app.post('/submit', async (req, res) => {
  const data = req.body;
  try {
    await appendDataToExcel(data);
    res.status(200).send({ message: 'Submission successful' });
  } catch (error) {
    console.error('Error saving data:', error);
    res.status(500).send({ message: 'Failed to save submission' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
