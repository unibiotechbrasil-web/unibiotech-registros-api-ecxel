import ExcelJS from 'exceljs';

// Reutilizar funções auxiliares do fabricados.js
function getAllKeysFromData(data) {
  const allKeys = new Set();
  for (const row of data) {
    Object.keys(row).forEach(key => allKeys.add(key));
  }
  return Array.from(allKeys);
}

function orderKeysDynamically(allKeys, preferredOrder) {
  const orderedKeys = [];
  for (const preferredKey of preferredOrder) {
    if (allKeys.includes(preferredKey)) {
      orderedKeys.push(preferredKey);
    }
  }
  
  const specialPatterns = ['ROT-', 'data_', 'lote_', 'total_', 'qtd_'];
  for (const pattern of specialPatterns) {
    const patternKeys = allKeys.filter(key => 
      key.startsWith(pattern) && !orderedKeys.includes(key)
    );
    patternKeys.sort();
    orderedKeys.push(...patternKeys);
  }
  
  const remainingKeys = allKeys.filter(key => !orderedKeys.includes(key));
  remainingKeys.sort();
  orderedKeys.push(...remainingKeys);
  
  return orderedKeys;
}

function formatHeaderName(key) {
  if (!key) return 'Coluna';
  return key
    .replace(/_/g, ' ')
    .replace(/-/g, ' ')
    .replace(/rot/gi, 'ROT')
    .split(' ')
    .map(word => {
      if (word.toUpperCase().startsWith('ROT')) return word.toUpperCase();
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ')
    .trim();
}

function isNumeric(value) {
  if (value === null || value === undefined) return false;
  return !isNaN(value) && !isNaN(parseFloat(value));
}

function isDateString(value) {
  if (typeof value !== 'string') return false;
  const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
  return dateRegex.test(value);
}

function isTimeString(value) {
  if (typeof value !== 'string') return false;
  const timeRegex = /^\d{2}:\d{2}(:\d{2})?(AM|PM|am|pm)?$/i;
  return timeRegex.test(value);
}

function filterByDate(data, dateField, startDate, endDate) {
  if ((!startDate || startDate.trim() === '') && (!endDate || endDate.trim() === '')) {
    return data;
  }
  
  return data.filter(item => {
    const dateStr = item[dateField];
    if (!dateStr || typeof dateStr !== 'string') return true;
    
    try {
      const [day, month, year] = dateStr.split('-').map(Number);
      const itemDate = new Date(year, month - 1, day);
      
      if (startDate) {
        const [sDay, sMonth, sYear] = startDate.split('-').map(Number);
        const start = new Date(sYear, sMonth - 1, sDay);
        if (itemDate < start) return false;
      }
      
      if (endDate) {
        const [eDay, eMonth, eYear] = endDate.split('-').map(Number);
        const end = new Date(eYear, eMonth - 1, eDay);
        if (itemDate > end) return false;
      }
      
      return true;
    } catch {
      return true;
    }
  });
}

// Handler principal
export default async function handler(req, res) {
  // Configurar CORS
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido. Use POST.' });
  }
  
  try {
    const { 
      data = [], 
      startDate = null, 
      endDate = null,
      fileName = `transportados_${Date.now()}`
    } = req.body;
    
    if (!Array.isArray(data)) {
      return res.status(400).json({ 
        success: false,
        error: 'Formato inválido. "data" deve ser um array de objetos.' 
      });
    }
    
    if (data.length === 0) {
      return res.status(400).json({ 
        success: false,
        error: 'Nenhum dado fornecido para exportação.' 
      });
    }
    
    // Filtrar por data
    const filteredData = filterByDate(data, 'data', startDate, endDate);
    
    if (filteredData.length === 0) {
      return res.status(404).json({ 
        success: false,
        error: 'Nenhum dado encontrado para o período selecionado.' 
      });
    }
    
    // Criar workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'UniBiotech App';
    workbook.created = new Date();
    
    const worksheet = workbook.addWorksheet('Transportados');
    
    // Configurar layout
    worksheet.pageSetup = {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    };
    
    // Título
    const titleRow = worksheet.getRow(1);
    const titleCell = titleRow.getCell(1);
    titleCell.value = 'RELATÓRIO DE TRANSPORTADOS';
    titleCell.font = { 
      name: 'Arial', 
      size: 18, 
      bold: true, 
      color: { argb: 'FF1F497D' } 
    };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:G1');
    
    // Obter e ordenar chaves
    const allKeys = getAllKeysFromData(filteredData);
    const preferredOrder = [
      'data', 
      'hora', 
      'cliente', 
      'transportadora', 
      'volume', 
      'organizador', 
      'levou', 
      'observação', 
      'criado_por', 
      'referência'
    ];
    
    const orderedKeys = orderKeysDynamically(allKeys, preferredOrder);
    const headers = orderedKeys.map(key => formatHeaderName(key));
    
    // Cabeçalho da tabela
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      
      cell.font = {
        name: 'Arial',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF808080' }
      };
      cell.border = {
        top: { style: 'medium' },
        bottom: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' }
      };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };
    });
    
    // Preencher dados
    filteredData.forEach((rowData, rowIndex) => {
      const dataRow = worksheet.getRow(rowIndex + 4);
      
      orderedKeys.forEach((key, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1);
        const value = rowData[key];
        
        if (value !== undefined && value !== null) {
          // Formatar hora
          if (key === 'hora' && isTimeString(value.toString())) {
            cell.value = value.toString();
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
          // Formatar números
          else if (isNumeric(value)) {
            cell.value = parseFloat(value);
            cell.numFmt = '#,##0';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          }
          // Formatar datas
          else if (isDateString(value.toString())) {
            cell.value = value.toString();
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          }
          // Texto normal
          else {
            cell.value = value.toString();
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
          }
        } else {
          cell.value = '';
        }
        
        // Bordas
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
    
    // Ajustar largura das colunas
    worksheet.columns.forEach((column, index) => {
      let maxLength = 0;
      worksheet.getColumn(index + 1).eachCell({ includeEmpty: true }, (cell) => {
        const length = cell.value ? cell.value.toString().length : 0;
        if (length > maxLength) {
          maxLength = length;
        }
      });
      column.width = Math.min(maxLength + 2, 30);
    });
    
    // Adicionar nota
    const lastRow = filteredData.length + 4;
    const noteRow = worksheet.getRow(lastRow + 2);
    const noteCell = noteRow.getCell(1);
    noteCell.value = 'NOTA: Este relatório inclui automaticamente todas as colunas encontradas nos dados.';
    noteCell.font = {
      name: 'Arial',
      italic: true,
      color: { argb: 'FF808080' }
    };
    
    if (orderedKeys.length > 1) {
      worksheet.mergeCells(`A${lastRow + 2}:${String.fromCharCode(65 + Math.min(orderedKeys.length - 1, 10))}${lastRow + 2}`);
    }
    
    // Gerar arquivo
    const buffer = await workbook.xlsx.writeBuffer();
    const base64Data = Buffer.from(buffer).toString('base64');
    
    return res.json({
      success: true,
      message: 'Planilha de transportados gerada com sucesso!',
      data: {
        fileName: `${fileName}.xlsx`,
        fileData: base64Data,
        fileType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: buffer.length,
        rows: filteredData.length,
        columns: orderedKeys.length,
        generatedAt: new Date().toISOString()
      }
    });
    
  } catch (error) {
    console.error('Erro ao gerar planilha:', error);
    return res.status(500).json({
      success: false,
      error: 'Erro interno ao gerar planilha',
      details: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
}
