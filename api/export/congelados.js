import ExcelJS from 'exceljs';

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
      fileName = `congelados_${Date.now()}`
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
    
    // Filtrar por data simples
    const filteredData = data.filter(item => {
      if (!startDate && !endDate) return true;
      
      const dateStr = item.data;
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
    
    // Separar por tipo
    const tipos = {};
    filteredData.forEach(item => {
      const tipo = item.tipo || 'GERAL';
      if (!tipos[tipo]) tipos[tipo] = [];
      tipos[tipo].push(item);
    });
    
    // Criar uma aba para cada tipo
    Object.entries(tipos).forEach(([tipo, items]) => {
      const safeSheetName = tipo.substring(0, 31).replace(/[\\/*?:\[\]]/g, '_');
      const worksheet = workbook.addWorksheet(safeSheetName);
      
      // Título da aba
      const titleRow = worksheet.getRow(1);
      const titleCell = titleRow.getCell(1);
      titleCell.value = `CONGELADOS - ${tipo}`;
      titleCell.font = { 
        name: 'Arial', 
        size: 14, 
        bold: true, 
        color: { argb: 'FF1F497D' } 
      };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.mergeCells('A1:E1');
      
      if (items.length > 0) {
        // Obter todas as chaves
        const allKeys = new Set();
        items.forEach(item => {
          Object.keys(item).forEach(key => allKeys.add(key));
        });
        
        // Ordenar chaves
        const preferredOrder = ['data', 'cliente', 'tipo', 'total'];
        const orderedKeys = [];
        
        preferredOrder.forEach(key => {
          if (allKeys.has(key)) {
            orderedKeys.push(key);
            allKeys.delete(key);
          }
        });
        
        // Adicionar ROT- primeiro
        const rotKeys = Array.from(allKeys).filter(key => key.startsWith('ROT-'));
        rotKeys.sort();
        orderedKeys.push(...rotKeys);
        
        // Restante
        const remainingKeys = Array.from(allKeys).filter(key => !key.startsWith('ROT-'));
        remainingKeys.sort();
        orderedKeys.push(...remainingKeys);
        
        // Cabeçalho
        const headerRow = worksheet.getRow(3);
        orderedKeys.forEach((key, index) => {
          const cell = headerRow.getCell(index + 1);
          const headerName = key
            .replace(/_/g, ' ')
            .split(' ')
            .map(word => {
              if (word.toUpperCase().startsWith('ROT')) return word.toUpperCase();
              return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
            })
            .join(' ');
          
          cell.value = headerName;
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
        
        // Dados
        items.forEach((item, rowIndex) => {
          const dataRow = worksheet.getRow(rowIndex + 4);
          
          orderedKeys.forEach((key, colIndex) => {
            const cell = dataRow.getCell(colIndex + 1);
            const value = item[key];
            
            if (value !== undefined && value !== null) {
              // Formatar números (incluindo ROT-)
              if (typeof value === 'number' || !isNaN(value)) {
                cell.value = parseFloat(value);
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
                
                // Destacar total
                if (key === 'total') {
                  cell.font = { ...cell.font, bold: true };
                }
              }
              // Formatar datas
              else if (typeof value === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(value)) {
                cell.value = value;
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
        
        // Ajustar colunas
        worksheet.columns.forEach((column, index) => {
          let maxLength = 0;
          worksheet.getColumn(index + 1).eachCell({ includeEmpty: true }, (cell) => {
            const length = cell.value ? cell.value.toString().length : 0;
            if (length > maxLength) {
              maxLength = length;
            }
          });
          column.width = Math.min(maxLength + 2, 20);
        });
      }
    });
    
    // Gerar arquivo
    const buffer = await workbook.xlsx.writeBuffer();
    const base64Data = Buffer.from(buffer).toString('base64');
    
    return res.json({
      success: true,
      message: 'Planilha de congelados gerada com sucesso!',
      data: {
        fileName: `${fileName}.xlsx`,
        fileData: base64Data,
        fileType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: buffer.length,
        rows: filteredData.length,
        sheets: Object.keys(tipos).length,
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
