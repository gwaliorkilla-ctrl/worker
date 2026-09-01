const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');

const app = express();

// Enable CORS & JSON Body Parsing
app.use(cors({ origin: '*' }));
app.use(express.json());

// Root test endpoint
app.get('/api', (req, res) => {
  res.status(200).send('Vercel API is live!');
});

// Primary generation endpoint
app.post('/api/generate', async (req, res) => {
  try {
    const { headers, rules } = req.body || {};

    if (!headers || headers.length === 0) {
      return res.status(400).json({ error: "No headers provided" });
    }

    const workbook = new ExcelJS.Workbook();
    const dataSheet = workbook.addWorksheet('Form_Template');
    const optionsSheet = workbook.addWorksheet('Options');

    // Write Header Row
    dataSheet.addRow(headers);

    let optColIndex = 1;
    headers.forEach((header, colIdx) => {
      const choices = rules[header];
      const colNum = colIdx + 1;

      if (choices && Array.isArray(choices) && choices.length > 0) {
        choices.forEach((choice, rowIdx) => {
          optionsSheet.getCell(rowIdx + 1, optColIndex).value = String(choice);
        });

        const startCell = optionsSheet.getCell(1, optColIndex).address;
        const endCell = optionsSheet.getCell(choices.length, optColIndex).address;
        const listFormula = `Options!$${startCell}:$${endCell}`;

        for (let r = 2; r <= 100; r++) {
          dataSheet.getCell(r, colNum).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [listFormula]
          };
        }
        optColIndex++;
      }
    });

    optionsSheet.state = 'hidden';

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Form_Template.xlsx"');
    return res.status(200).send(Buffer.from(buffer));

  } catch (err) {
    return res.status(500).json({ error: err.message, stack: err.stack });
  }
});

// Catch-all for Express on Vercel
module.exports = app;
