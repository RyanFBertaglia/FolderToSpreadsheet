import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

const sheetPath = './planilha.xlsx';
const folderPath = './images';
const aba = 'fotos';
const coluna = 1; // 1 é a coluna A, 2 é a coluna B, etc.

async function main() {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(sheetPath);
  } catch (error) {
    workbook.addWorksheet(aba);
  }

  let sheet = workbook.getWorksheet(aba);
  if (!sheet) sheet = workbook.addWorksheet(aba);

  const images = fs.readdirSync(folderPath).filter(file =>
    file.endsWith('.png') || file.endsWith('.jpg') || file.endsWith('.jpeg')
  );

  for (let i = 0; i < images.length; i++) {
    const image = images[i];
    const imagePath = path.join(folderPath, image);

    const imageId = workbook.addImage({
      filename: imagePath,
      extension: path.extname(image).substring(1),
    });

    sheet.addImage(imageId, {
      tl: { col: coluna - 1, row: i * 5 }, // espaçamento de 5 linhas
      br: { col: coluna, row: (i * 5) + 5 },
    });
  }

  await workbook.xlsx.writeFile(sheetPath);

  for (const image of images) {
    const imagePath = path.join(folderPath, image);
    fs.unlinkSync(imagePath);
  }

  console.log('Imagens inseridas');
}

main();
