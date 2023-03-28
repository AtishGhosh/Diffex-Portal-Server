import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { convertToIndianWords } from './number-to-rupees.js';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default function generateInvoice(req, res) {
    const {
    invoiceNumber,
    invoiceDate,
    poNumber,
    poDate,
    productionOrder,
    placeOfSupply,
    receiverName,
    receiverAddress,
    receiverGSTIN,
    receiverState,
    receiverStateCode,
    shippedName,
    shippedAddress,
    shippedGSTIN,
    shippedState,
    shippedStateCode,
    bankDetails,
    accountNumber,
    bankBranchIFSC,
    srNo,
    nameOfProduct,
    hsn,
    quantity,
    rate,
    amount,
    discount,
    discountAmount,
    taxableValue,
    cgst,
    cgstAmount,
    sgst,
    sgstAmount,
    igst,
    igstAmount,
    totalValue,
    grandTotalBeforeTax,
    grandTotalCGSTamount,
    grandTotalSGSTamount,
    grandTotalIGSTamount,
    grandTotalGSTamount,
    grandTotalAfterTax
  } = req.body;
    
  // Your code to generate the invoice goes here
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Different Experience Enterprises';
  const worksheet = workbook.addWorksheet('Invoice');
    
  const defaultRowHeight = 25;
    
  worksheet.getColumn('A').width = 4.5;
  worksheet.getColumn('B').width = 16.5;
  worksheet.getColumn('C').width = 9.5;
  worksheet.getColumn('D').width = 4.5;
  worksheet.getColumn('E').width = 4.5;
  worksheet.getColumn('F').width = 10.5;
  worksheet.getColumn('G').width = 10.5;
  worksheet.getColumn('H').width = 7.5;
  worksheet.getColumn('I').width = 10.5;
  worksheet.getColumn('J').width = 5.5;
  worksheet.getColumn('K').width = 8.5;
  worksheet.getColumn('L').width = 5.5;
  worksheet.getColumn('M').width = 8.5;
  worksheet.getColumn('N').width = 5.5;
  worksheet.getColumn('O').width = 8.5;
  worksheet.getColumn('P').width = 10.5;
    
  const columnNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'];
    
  const defaultCellBorder = {
    top: {style: 'thin',},
    left: {style: 'thin',},
    bottom: {style: 'thin',},
    right: {style: 'thin',},
  };
    
  var cell;
    
  worksheet.mergeCells('A1:P2');
  cell = worksheet.getCell('A1');
  cell.value = 'Different Experience Enterprises';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 24,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A3:P4');
  cell = worksheet.getCell('A3');
  cell.value = 'Flat No. B1, Royale Apartment, 920, Madurdaha, Kolkata - 700107';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A5:K7');
  cell = worksheet.getCell('A5');
  cell.value = 'TAX INVOICE';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 24,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  cell = worksheet.getCell('L5'); 
  cell.style = {
    border: defaultCellBorder,
  };
    
  worksheet.mergeCells('M5:P5');
  cell = worksheet.getCell('M5');
  cell.value = 'Original for Receipient';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  cell = worksheet.getCell('L6');
  cell.style = {
    border: defaultCellBorder,
  };
    
  worksheet.mergeCells('M6:P6');
  cell = worksheet.getCell('M6');
  cell.value = 'Duplicate for Supplier/Transporter';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  cell = worksheet.getCell('L7');
  cell.style = {
    border: defaultCellBorder,
  };
    
  worksheet.mergeCells('M7:P7');
  cell = worksheet.getCell('M7');
  cell.value = 'Triplicate for Supplier';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Invoice Details
    
  worksheet.mergeCells('A8:B8');
  cell = worksheet.getCell('A8');
  cell.value = 'Our GST No. :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A9:B9');
  cell = worksheet.getCell('A9');
  cell.value = 'Invoice No. :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A10:B10');
  cell = worksheet.getCell('A10');
  cell.value = 'Invoice Date :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A11:B11');
  cell = worksheet.getCell('A11');
  cell.value = 'State :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('C8:I8');
  cell = worksheet.getCell('C8');
  cell.value = '19ACAPG5401B1ZD';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Invoice Number
  worksheet.mergeCells('C9:I9');
  cell = worksheet.getCell('C9');
  cell.value = invoiceNumber.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 14,
      bold: true,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Invoice Date
  worksheet.mergeCells('C10:I10');
  cell = worksheet.getCell('C10');
  cell.value = invoiceDate.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('C11:E11');
  cell = worksheet.getCell('C11');
  cell.value = 'West Bengal';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('F11:G11');
  cell = worksheet.getCell('F11');
  cell.value = 'State Code :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('H11:I11');
  cell = worksheet.getCell('H11');
  cell.value = '19';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J8:L8');
  cell = worksheet.getCell('J8');
  cell.value = 'PO No. :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J9:L9');
  cell = worksheet.getCell('J9');
  cell.value = 'PO Date :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J10:L10');
  cell = worksheet.getCell('J10');
  cell.value = 'Production Order :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J11:L11');
  cell = worksheet.getCell('J11');
  cell.value = 'Place of Supply :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // PO Number
  worksheet.mergeCells('M8:P8');
  cell = worksheet.getCell('M8');
  cell.value = poNumber.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Current Date
  worksheet.mergeCells('M9:P9');
  cell = worksheet.getCell('M9');
  cell.value = poDate.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Production Order
  worksheet.mergeCells('M10:P10');
  if (productionOrder !== '') {
    cell = worksheet.getCell('M10');
    cell.value = productionOrder.toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
    name: 'Arial',
    size: 12,
      },
      alignment: {
    vertical: 'middle',
    wrapText: true,
      },
    };
  }
    
  // Place of Supply
  worksheet.mergeCells('M11:P11');
  cell = worksheet.getCell('M11');
  cell.value = placeOfSupply.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Details of Receiver / Billed to
    
  worksheet.mergeCells('A12:I12');
  cell = worksheet.getCell('A12');
  cell.value = 'Details of Receiver / Billed to :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
      bold: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A13:B13');
  cell = worksheet.getCell('A13');
  cell.value = 'Name :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A14:B14');
  cell = worksheet.getCell('A14');
  cell.value = 'Address :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A15:B15');
  cell = worksheet.getCell('A15');
  cell.value = 'GSTIN :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('A16:B16');
  cell = worksheet.getCell('A16');
  cell.value = 'State :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Receiver Name
  worksheet.mergeCells('C13:I13');
  cell = worksheet.getCell('C13');
  cell.value = receiverName.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Receiver Address
  worksheet.mergeCells('C14:I14');
  cell = worksheet.getCell('C14');
  cell.value = receiverAddress.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Receiver GSTIN
  worksheet.mergeCells('C15:I15');
  cell = worksheet.getCell('C15');
  cell.value = receiverGSTIN.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Receiver State
  worksheet.mergeCells('C16:E16');
  cell = worksheet.getCell('C16');
  cell.value = receiverState.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('F16:G16');
  cell = worksheet.getCell('F16');
  cell.value = 'State Code :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Receiver State Code
  worksheet.mergeCells('H16:I16');
  cell = worksheet.getCell('H16');
  cell.value = receiverStateCode.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Details of Consignee / Shipped to
    
  worksheet.mergeCells('J12:P12');
  cell = worksheet.getCell('J12');
  cell.value = 'Details of Consignee / Shipped to :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
      bold: true,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J13:K13');
  cell = worksheet.getCell('J13');
  cell.value = 'Name :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J14:K14');
  cell = worksheet.getCell('J14');
  cell.value = 'Address :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J15:K15');
  cell = worksheet.getCell('J15');
  cell.value = 'GSTIN :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J16:K16');
  cell = worksheet.getCell('K16');
  cell.value = 'State :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Shipped Name
  worksheet.mergeCells('L13:P13');
  cell = worksheet.getCell('L13');
  cell.value = shippedName.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Shipped Address
  worksheet.mergeCells('L14:P14');
  cell = worksheet.getCell('L14');
  cell.value = shippedAddress.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Shipped GSTIN
  worksheet.mergeCells('L15:P15');
  cell = worksheet.getCell('L15');
  cell.value = shippedGSTIN.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Shipped State
  worksheet.mergeCells('L16:M16');
  cell = worksheet.getCell('L16');
  cell.value = shippedState.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('N16:O16');
  cell = worksheet.getCell('N16');
  cell.value = 'State Code :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Shipped State Code
  // worksheet.mergeCells('P16:P16');
  cell = worksheet.getCell('P16');
  cell.value = shippedStateCode.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 12,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Products
    
  // Sr. No. Header
  worksheet.mergeCells('A18:A19');
  cell = worksheet.getCell('A18');
  cell.value = 'Sr No';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Name of Product / Service Header
  worksheet.mergeCells('B18:B19');
  cell = worksheet.getCell('B18');
  cell.value = 'Name of Product / Service';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // HSN Header
  worksheet.mergeCells('C18:C19');
  cell = worksheet.getCell('C18');
  cell.value = 'HSN / ASN';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // UMO Header
  worksheet.mergeCells('D18:D19');
  cell = worksheet.getCell('D18');
  cell.value = 'UMO';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Qty Header
  worksheet.mergeCells('E18:E19');
  cell = worksheet.getCell('E18');
  cell.value = 'Qty';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Rate Header
  worksheet.mergeCells('F18:F19');
  cell = worksheet.getCell('F18');
  cell.value = 'Rate';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Amount Header
  worksheet.mergeCells('G18:G19');
  cell = worksheet.getCell('G18');
  cell.value = 'Amount';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Discount Header
  worksheet.mergeCells('H18:H19');
  cell = worksheet.getCell('H18');
  cell.value = 'Less : Dis';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Taxable Value Header
  worksheet.mergeCells('I18:I19');
  cell = worksheet.getCell('I18');
  cell.value = 'Taxable Value';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // CGST Header
  worksheet.mergeCells('J18:K18');
  cell = worksheet.getCell('J18');
  cell.value = 'CGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('J19:J19');
  cell = worksheet.getCell('J19');
  cell.value = 'Rate';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('K19:K19');
  cell = worksheet.getCell('K19');
  cell.value = 'Amount';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // SGST Header
  worksheet.mergeCells('L18:M18');
  cell = worksheet.getCell('L18');
  cell.value = 'SGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('L19:L19');
  cell = worksheet.getCell('L19');
  cell.value = 'Rate';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('M19:M19');
  cell = worksheet.getCell('M19');
  cell.value = 'Amount';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // IGST Header
  worksheet.mergeCells('N18:O18');
  cell = worksheet.getCell('N18');
  cell.value = 'IGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('N19:N19');
  cell = worksheet.getCell('N19');
  cell.value = 'Rate';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // worksheet.mergeCells('O19:O19');
  cell = worksheet.getCell('O19');
  cell.value = 'Amount';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  // Total Amount Header
  worksheet.mergeCells('P18:P19');
  cell = worksheet.getCell('P18');
  cell.value = 'Total Amount';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  var currentRow = 20;
    
  for (var i = 1; i < currentRow; i++) {
    if (i === 14) {
      worksheet.getRow(i).height = 50;
      continue;
    }
    worksheet.getRow(i).height = defaultRowHeight;
  }
    
  for (var i = 0; i < srNo.length; i++) {
    var currentRowValue = currentRow.toString();
    const row = worksheet.getRow(currentRow);
    row.height = 100;
    
    // Sr. No. Value
    cell = worksheet.getCell('A' + currentRowValue);
    cell.value = srNo[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Name of Product / Service Value
    cell = worksheet.getCell('B' + currentRowValue);
    cell.value = nameOfProduct[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // HSN Value
    cell = worksheet.getCell('C' + currentRowValue);
    cell.value = hsn[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // UMO Value
    cell = worksheet.getCell('D' + currentRowValue);
    cell.value = 'NOS';
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Quantity Value
    cell = worksheet.getCell('E' + currentRowValue);
    cell.value = quantity[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Rate Value
    cell = worksheet.getCell('F' + currentRowValue);
    cell.value = rate[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Amount Value
    cell = worksheet.getCell('G' + currentRowValue);
    cell.value = amount[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Discount Value
    cell = worksheet.getCell('H' + currentRowValue);
    if (discount[i].toString() === '' || discount[i] === 0 || discount[i].toString() === '0' || discountAmount[i].toString() === '' || discountAmount[i] === 0 || discountAmount[i].toString() === '0') {
      cell.value = 'NIL';
    } else {
      cell.value = discount[i].toString() + '% : ' + discountAmount[i].toString();
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Taxable Value
    cell = worksheet.getCell('I' + currentRowValue);
    cell.value = taxableValue[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // CGST Rate Value
    cell = worksheet.getCell('J' + currentRowValue);
    if (cgst[i].toString() === '' || cgst[i] === 0 || cgst[i].toString() === '0') {
      cell.value = '0%';
    } else {
      cell.value = cgst[i].toString() + "%";
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // CGST Amount Value
    cell = worksheet.getCell('K' + currentRowValue);
    if (cgstAmount[i].toString() === '' || cgstAmount[i] === 0 || cgstAmount[i].toString() === '0') {
      cell.value = '0.00';
    } else {
      cell.value = cgstAmount[i].toString();
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // SGST Rate Value
    cell = worksheet.getCell('L' + currentRowValue);
    if (sgst[i].toString() === '' || sgst[i] === 0 || sgst[i].toString() === '0') {
      cell.value = '0%';
    } else {
      cell.value = sgst[i].toString() + "%";
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // SGST Amount Value
    cell = worksheet.getCell('M' + currentRowValue);
    if (sgstAmount[i].toString() === '' || sgstAmount[i] === 0 || sgstAmount[i].toString() === '0') {
      cell.value = '0.00';
    } else {
      cell.value = sgstAmount[i].toString();
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // IGST Rate Value
    cell = worksheet.getCell('N' + currentRowValue);
    if (igst[i].toString() === '' || igst[i] === 0 || igst[i].toString() === '0') {
      cell.value = '0%';
    } else {
      cell.value = igst[i].toString() + "%";
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // IGST Amount Value
    cell = worksheet.getCell('O' + currentRowValue);
    if (igstAmount[i].toString() === '' || igstAmount[i] === 0 || igstAmount[i].toString() === '0') {
      cell.value = '0.00';
    } else {
      cell.value = igstAmount[i].toString();
    }
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    // Total Amount
    cell = worksheet.getCell('P' + currentRowValue);
    cell.value = totalValue[i].toString();
    cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
    currentRow += 1;
  }
    
  const rowAfterProducts = currentRow;
    
  // Skipping a row
  currentRow += 1;
    
  // Total Amount Header
  worksheet.mergeCells('A' + currentRow.toString() + ':D' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = 'Total';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  cell = worksheet.getCell('I' + currentRow.toString());
  cell.value = grandTotalBeforeTax.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':K' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = grandTotalCGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('L' + currentRow.toString() + ':M' + currentRow.toString());
  cell = worksheet.getCell('L' + currentRow.toString());
  cell.value = grandTotalSGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('N' + currentRow.toString() + ':O' + currentRow.toString());
  cell = worksheet.getCell('N' + currentRow.toString());
  cell.value = grandTotalIGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  cell = worksheet.getCell('P' + currentRow.toString());
  cell.value = grandTotalAfterTax.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':I' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = 'Total Invoice Amount in Words :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 11,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Total Amount Before Tax';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalBeforeTax.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':I' + (currentRow + 3).toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = convertToIndianWords(grandTotalAfterTax.toString()).toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 11,
    },
    alignment: {
      vertical: 'top',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Add : CGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalCGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Add : SGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalSGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Add : IGST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalIGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Tax Amount : GST';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalGSTamount.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':F' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = 'Bank Details : ' + bankDetails.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('G' + currentRow.toString() + ':I' + (currentRow + 8).toString());
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Total Amount After Tax';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = grandTotalAfterTax.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':B' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = 'Bank A/C No. :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('C' + currentRow.toString() + ':F' + currentRow.toString());
  cell = worksheet.getCell('C' + currentRow.toString());
  cell.value = accountNumber.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':P' + currentRow.toString());
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':B' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = 'Bank Branch IFSC :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('C' + currentRow.toString() + ':F' + currentRow.toString());
  cell = worksheet.getCell('C' + currentRow.toString());
  cell.value = bankBranchIFSC.toString();
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':N' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'GST Payable on Reverse Charge';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('O' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('O' + currentRow.toString());
  cell.value = 'N.A';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'right',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':F' + currentRow.toString());
  cell = worksheet.getCell('A' + currentRow.toString());
  cell.value = ': Remarks :';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Certified that the particulars given above are true and correct.';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 8,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('A' + currentRow.toString() + ':F' + (currentRow + 5).toString());
    
  worksheet.mergeCells('J' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'For, Different Experience Enterprises';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  currentRow += 1;
    
  worksheet.mergeCells('J' + currentRow.toString() + ':P' + (currentRow + 3).toString());
    
  currentRow += 4;
    
  worksheet.mergeCells('G' + currentRow.toString() + ':I' + currentRow.toString());
  cell = worksheet.getCell('G' + currentRow.toString());
  cell.value = '(Company Seal)';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  worksheet.mergeCells('J' + currentRow.toString() + ':P' + currentRow.toString());
  cell = worksheet.getCell('J' + currentRow.toString());
  cell.value = 'Authorised Signatory';
  cell.style = {
    border: defaultCellBorder,
    font: {
      name: 'Arial',
      size: 10,
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    },
  };
    
  for (var i = rowAfterProducts; i <= currentRow; i++) {
    worksheet.getRow(i).height = defaultRowHeight;
  }
    
  for (var i = 1; i <= currentRow; i++) {
    columnNames.forEach(columnName => {
      const eachCell = worksheet.getCell(columnName.toString() + i.toString());
      if (eachCell.value === null || eachCell.value === undefined || eachCell.value === '') {
    eachCell.style.border = defaultCellBorder;
      }
      var eachCellBefore;
      if (columnName !== 'A') {
    eachCellBefore = worksheet.getCell(columnNames[columnNames.indexOf(columnName.toString()) - 1].toString() + i.toString());
    if (eachCellBefore.value === eachCell.value) {
      eachCell.style.border = defaultCellBorder;
    }
      }
      if (i > 1) {
    eachCellBefore = worksheet.getCell(columnName.toString() + (i - 1).toString());
    if (eachCellBefore.value === eachCell.value) {
      eachCell.style.border = defaultCellBorder;
    }
      }
    });
  }
  
  // Set the content type and disposition of the response
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', ('attachment; filename="Invoice ' + invoiceNumber.toString() + '.xlsx"'));

  const newInvoiceDataJSON = {
    invoiceNumber,
    invoiceDate,
    poNumber,
    poDate,
    productionOrder,
    placeOfSupply,
    receiverName,
    receiverAddress,
    receiverGSTIN,
    receiverState,
    receiverStateCode,
    shippedName,
    shippedAddress,
    shippedGSTIN,
    shippedState,
    shippedStateCode,
    bankDetails,
    accountNumber,
    bankBranchIFSC,
    srNo,
    nameOfProduct,
    hsn,
    quantity,
    rate,
    amount,
    discount,
    discountAmount,
    taxableValue,
    cgst,
    cgstAmount,
    sgst,
    sgstAmount,
    igst,
    igstAmount,
    totalValue,
    grandTotalBeforeTax,
    grandTotalCGSTamount,
    grandTotalSGSTamount,
    grandTotalIGSTamount,
    grandTotalGSTamount,
    grandTotalAfterTax
  };  

  // Write the data to the JSON file
  const filePath = path.join(__dirname, '../storage/data/invoices.json');
  const jsonData = fs.readFileSync(filePath);
  const invoices = JSON.parse(jsonData);
  invoices.push(newInvoiceDataJSON);
  fs.writeFileSync(filePath, JSON.stringify(invoices, null, 2));

  // Send the workbook as a buffer
  workbook.xlsx.write(res)
    .then(() => {
      res.status(200);
      res.end();
    })
    .catch((err) => {
      console.log(err);
      res.status(500).send('An error occurred while generating the Excel file.');
    });
}