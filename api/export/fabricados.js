const ExcelJS = require('exceljs');

// Função auxiliar para obter todas as chaves
function getAllKeysFromData(data) {
  const allKeys = new Set();
  for (const row of data) {
    Object.keys(row).forEach(key => allKeys.add(key));
  }
  return Array.from(allKeys);
}

// Função para ordenar chaves
function orderKeysDynamically(allKeys, preferredOrder) {
  const orderedKeys = [];
  
  // 1. Adicionar chaves preferenciais
  for (const preferredKey of preferredOrder) {
    if (allKeys.includes(preferredKey)) {
      orderedKeys.push(preferredKey);
    }
  }
  
  // 2. Adicionar chaves especiais
  const specialPatterns = ['ROT-', 'data_', 'lote_', 'total_', 'qtd_'];
  for (const pattern of specialPatterns) {
    const patternKeys = allKeys.filter(key => 
      key.startsWith(pattern) && !orderedKeys.includes(key)
    );
    patternKeys.sort();
    orderedKeys.push(...patternKeys);
  }
  
  // 3. Restante em ordem alfabética
  const remainingKeys = allKeys.filter(key => !orderedKeys.includes(key));
  remainingKeys.sort();
  orderedKeys.push(...remainingKeys);
  
  return orderedKeys;
}

// Formatar nome do cabeçalho
function formatHeaderName(key) {
  if (!key) return 'Coluna';
  
  return key
    .replace(/_/g, ' ')
    .replace(/-/g, ' ')
    .replace(/rot/gi, 'ROT')
    .split(' ')
    .map(word => {
      if (word.toUpperCase().startsWith('ROT')) {
        return word.toUpperCase();
      }
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ')
    .trim();
}

// Verificar se é numérico
function isNumeric(value) {
  if (value === null || value === undefined) return false;
  return !isNaN(value) && !isNaN(parseFloat(value));
}

// Verificar se é data no formato dd-MM-yyyy
function isDateString(value) {
  if (typeof value !== 'string') return false;
  const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
  return dateRegex.test(value);
}

// Filtrar por data
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
module.exports = async function handler(req, res) {
  // Configurar CORS
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  
  // Responder preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  // Apenas POST permitido
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido. Use POST.' });
  }
  
  try {
    const { 
      data = [], 
      startDate = null, 
      endDate = null,
      fileName = fabricados_${Date.now()}
    } = req.body;
    
    // Validar dados
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
    
    // Filtrar por data se necessário
    const filteredData = filterByDate(data, 'data', startDate, endDate);
    
    if (filteredData.length === 0) {
      return res.status(404).json({ 
        success: false,
        error: 'Nenhum dado encontrado para o período selecionado.' 
      });
    }
    
    // Criar workbook Excel
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'UniBiotech App';
    workbook.created = new Date();
    
    // Criar worksheet
    const worksheet = workbook.addWorksheet('Fabricados');
    
    // Configurar layout da página
    worksheet.pageSetup = {
      paperSize: 9, // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    };
    
    // Adicionar título
    const titleRow = worksheet.getRow(1);
    const titleCell = titleRow.getCell(1);
    titleCell.value = 'RELATÓRIO DE FABRICADOS';
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
      'nome_produto', 
      'lote_biotech', 
      'lote_produto', 
      'quantidade', 
      'data_fabricação', 
      'data_validade', 
      'criado', 
      'observação', 
      'referência'
    ];
    
    const orderedKeys = orderKeysDynamically(allKeys, preferredOrder);
    const headers = orderedKeys.map(key => formatHeaderName(key));
    
    // Criar cabeçalho da tabela (linha 3)
    const headerRow = worksheet.getRow(3);
    headers.forEach((header, index) => {
      const cell = headerRow.getCell(index + 1);
      cell.value = header;
      
      // Estilo do cabeçalho
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
      const dataRow = worksheet.getRow(rowIndex + 4); // Começar na linha 4
      
      orderedKeys.forEach((key, colIndex) => {
        const cell = dataRow.getCell(colIndex + 1);
        const value = rowData[key];
        
        if (value !== undefined && value !== null) {
          // Formatar números
          if (isNumeric(value)) {
            // MANTÉM LOTES COMO TEXTO - SEM FORMATAÇÃO DE MILHAR
            if (key.toLowerCase().includes('lote') || 
                key.toLowerCase().includes('cod') || 
                key.toLowerCase().includes('ref')) {
              // Lotese códigos: mantém como texto exato
              cell.value = value.toString();
              cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } 
            // Quantidades: formata com separador de milhar
            else if (key.toLowerCase().includes('quantidade') || 
                     key.toLowerCase().includes('qtd')) {
              cell.value = parseFloat(value);
              cell.numFmt = '#,##0';
              cell.alignment = { horizontal: 'right', vertical: 'middle' };
            }
            // Outros números pequenos: mantém como texto
            else {
              const numValue = parseFloat(value);
              if (numValue <= 99999) {
                cell.value = value.toString();
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
              } else {
                cell.value = numValue;
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
              }
            }
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
        
        // Bordas para todas as células
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
    
    // Adicionar nota sobre colunas dinâmicas
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
      worksheet.mergeCells(A${lastRow + 2}:${String.fromCharCode(65 + Math.min(orderedKeys.length - 1, 10))}${lastRow + 2});
    }
    
    // Gerar buffer do Excel
    const buffer = await workbook.xlsx.writeBuffer();
    
    // Converter para base64
    const base64Data = Buffer.from(buffer).toString('base64');
    
    // Retornar resposta
    return res.json({
      success: true,
      message: 'Planilha de fabricados gerada com sucesso!',
      data: {
        fileName: ${fileName}.xlsx,
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