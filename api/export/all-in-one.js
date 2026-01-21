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
      fabricados = [], 
      transportados = [], 
      congelados = [],
      startDate = null, 
      endDate = null,
      fileName = `relatorio_completo_${Date.now()}`
    } = req.body;
    
    if (fabricados.length === 0 && transportados.length === 0 && congelados.length === 0) {
      return res.status(400).json({ 
        success: false,
        error: 'Nenhum dado fornecido para exportação.' 
      });
    }
    
    // Função simples de filtro por data
    const filterData = (data, dateField) => {
      if (!startDate && !endDate) return data;
      
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
    };
    
    const filteredFabricados = filterData(fabricados, 'data');
    const filteredTransportados = filterData(transportados, 'data');
    const filteredCongelados = filterData(congelados, 'data');
    
    // Criar workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'UniBiotech App';
    workbook.created = new Date();
    
    // Função para criar aba genérica
    const createGenericSheet = (data, sheetName, title) => {
      if (data.length === 0) return;
      
      const worksheet = workbook.addWorksheet(sheetName.substring(0, 31));
      
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
      titleCell.value = title;
      titleCell.font = { 
        name: 'Arial', 
        size: 16, 
        bold: true, 
        color: { argb: 'FF1F497D' } 
      };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.mergeCells('A1:G1');
      
      if (data.length > 0) {
        // Obter todas as chaves
        const allKeys = new Set();
        data.forEach(item => {
          Object.keys(item).forEach(key => allKeys.add(key));
        });
        
        const keysArray = Array.from(allKeys);
        
        // Cabeçalho
        const headerRow = worksheet.getRow(3);
        keysArray.forEach((key, index) => {
          const cell = headerRow.getCell(index + 1);
          const headerName = key
            .replace(/_/g, ' ')
            .split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
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
        data.forEach((item, rowIndex) => {
          const dataRow = worksheet.getRow(rowIndex + 4);
          
          keysArray.forEach((key, colIndex) => {
            const cell = dataRow.getCell(colIndex + 1);
            const value = item[key];
            
            if (value !== undefined && value !== null) {
              // Formatar números
              if (typeof value === 'number' || !isNaN(value)) {
                cell.value = parseFloat(value);
                cell.numFmt = '#,##0';
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
              }
              // Formatar datas (dd-mm-yyyy)
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
          column.width = Math.min(maxLength + 2, 25);
        });
      }
    };
    
    // Criar abas para cada tipo de dado
    if (filteredFabricados.length > 0) {
      createGenericSheet(filteredFabricados, 'FABRICADOS', 'RELATÓRIO DE FABRICADOS');
    }
    
    if (filteredTransportados.length > 0) {
      createGenericSheet(filteredTransportados, 'TRANSPORTADOS', 'RELATÓRIO DE TRANSPORTADOS');
    }
    
    if (filteredCongelados.length > 0) {
      // Para congelados, criar aba simplificada
      const worksheet = workbook.addWorksheet('CONGELADOS');
      
      // Título
      const titleRow = worksheet.getRow(1);
      const titleCell = titleRow.getCell(1);
      titleCell.value = 'RELATÓRIO DE CONGELADOS';
      titleCell.font = { 
        name: 'Arial', 
        size: 16, 
        bold: true, 
        color: { argb: 'FF1F497D' } 
      };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.mergeCells('A1:G1');
      
      if (filteredCongelados.length > 0) {
        // Obter chaves
        const allKeys = new Set();
        filteredCongelados.forEach(item => {
          Object.keys(item).forEach(key => allKeys.add(key));
        });
        
        const keysArray = Array.from(allKeys);
        
        // Cabeçalho
        const headerRow = worksheet.getRow(3);
        keysArray.forEach((key, index) => {
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
        filteredCongelados.forEach((item, rowIndex) => {
          const dataRow = worksheet.getRow(rowIndex + 4);
          
          keysArray.forEach((key, colIndex) => {
            const cell = dataRow.getCell(colIndex + 1);
            const value = item[key];
            
            if (value !== undefined && value !== null) {
              // Formatar números (ROT- e total)
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
          column.width = Math.min(maxLength + 2, 15);
        });
      }
    }
    
    // Gerar arquivo
    const buffer = await workbook.xlsx.writeBuffer();
    const base64Data = Buffer.from(buffer).toString('base64');
    
    return res.json({
      success: true,
      message: 'Relatório completo gerado com sucesso!',
      data: {
        fileName: `${fileName}.xlsx`,
        fileData: base64Data,
        fileType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: buffer.length,
        sheets: workbook.worksheets.length,
        summary: {
          fabricados: filteredFabricados.length,
          transportados: filteredTransportados.length,
          congelados: filteredCongelados.length,
          total: filteredFabricados.length + filteredTransportados.length + filteredCongelados.length
        },
        generatedAt: new Date().toISOString()
      }
    });
    
  } catch (error) {
    console.error('Erro ao gerar relatório completo:', error);
    return res.status(500).json({
      success: false,
      error: 'Erro interno ao gerar relatório',
      details: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
}
