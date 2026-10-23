const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');

const app = express();

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '50mb' }));

app.get('/api', (req, res) => {
  res.status(200).send('Vercel API is live!');
});

app.post('/api/generate', async (req, res) => {
  try {
    const { headers, rules = {}, meta = {} } = req.body || {};

    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      return res.status(400).json({ error: "No headers provided" });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Smart Form Excel Automator';
    workbook.created = new Date();

    const dataSheet = workbook.addWorksheet('Form_Template', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });
    const optionsSheet = workbook.addWorksheet('Options');

    // Add and style header row
    const headerRow = dataSheet.addRow(headers);
    headerRow.height = 28;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E78' } // Navy blue professional styling
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        bottom: { style: 'medium', color: { argb: 'FF102A43' } },
        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
      };
    });

    let optColIndex = 1;
    headers.forEach((header, colIdx) => {
      const colNum = colIdx + 1;
      const targetColLetter = dataSheet.getColumn(colNum).letter;
      const choices = (rules && rules[header]) || (meta && meta[header]?.choices);
      const fieldMeta = (meta && meta[header]) || {};

      let maxLen = String(header).length;

      // 1. Dropdown List Validation (Select, Radio, Checkbox options)
      if (choices && Array.isArray(choices) && choices.length > 0) {
        const optColLetter = optionsSheet.getColumn(optColIndex).letter;
        choices.forEach((choice, rowIdx) => {
          optionsSheet.getCell(rowIdx + 1, optColIndex).value = String(choice);
          if (String(choice).length > maxLen) maxLen = String(choice).length;
        });

        // Use absolute row & column with single-quoted sheet name
        const listFormula = `'Options'!$${optColLetter}$1:$${optColLetter}$${choices.length}`;

        dataSheet.dataValidations.add(`${targetColLetter}2:${targetColLetter}100`, {
          type: 'list',
          allowBlank: true,
          formulae: [listFormula],
          showErrorMessage: true,
          errorTitle: 'Invalid Selection',
          error: 'Please choose an option from the dropdown list.'
        });
        optColIndex++;
      } 
      // 2. Date Validation
      else if (fieldMeta.type === 'date') {
        dataSheet.dataValidations.add(`${targetColLetter}2:${targetColLetter}100`, {
          type: 'date',
          operator: 'between',
          allowBlank: true,
          formulae: [new Date(1900, 0, 1), new Date(2099, 11, 31)],
          showErrorMessage: true,
          errorTitle: 'Invalid Date',
          error: 'Please enter a valid date in YYYY-MM-DD format.'
        });
        dataSheet.getColumn(colNum).numFmt = 'yyyy-mm-dd';
      } 
      // 3. Whole Number / Numeric Validation
      else if (fieldMeta.type === 'number') {
        const min = fieldMeta.min !== undefined && fieldMeta.min !== '' ? Number(fieldMeta.min) : 0;
        const max = fieldMeta.max !== undefined && fieldMeta.max !== '' ? Number(fieldMeta.max) : 999999999;
        dataSheet.dataValidations.add(`${targetColLetter}2:${targetColLetter}100`, {
          type: 'whole',
          operator: 'between',
          allowBlank: true,
          formulae: [min, max],
          showErrorMessage: true,
          errorTitle: 'Invalid Number',
          error: `Please enter a whole number between ${min} and ${max}.`
        });
      }

      // Column Auto-width with reasonable bounds
      dataSheet.getColumn(colNum).width = Math.min(Math.max(maxLen + 4, 15), 45);
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

const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});

module.exports = app;
