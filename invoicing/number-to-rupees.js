function convertDigitsToWords(num) {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const tens = ['', 'Ten', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  let words = '';
  if (num < 10) {
    words = ones[num];
  }
  else if (num < 20) {
    words = teens[num - 10];
  }
  else if (num < 100) {
    words = `${tens[Math.floor(num / 10)]} ${ones[num % 10]}`;
  }
  else if (num < 1000) {
    words = `${ones[Math.floor(num / 100)]} Hundred ${convertDigitsToWords(num % 100)}`;
  }
  else if (num < 10000) {
    words = `${ones[Math.floor(num / 1000)]} Thousand ${convertDigitsToWords(num % 1000)}`;
  }
  else if (num < 100000) {
    words = `${convertDigitsToWords(Math.floor(num / 1000))} Thousand ${convertDigitsToWords(num % 1000)}`;
  }
  else if (num < 1000000) {
    words = `${ones[Math.floor(num / 100000)]} Lakh ${convertDigitsToWords(num % 100000)}`;
  }
  else if (num < 10000000) {
    words = `${convertDigitsToWords(Math.floor(num / 100000))} Lakh ${convertDigitsToWords(num % 100000)}`;
  }
  else if (num < 100000000) {
    words = `${ones[Math.floor(num / 10000000)]} Crore ${convertDigitsToWords(num % 10000000)}`;
  }
  else if (num < 1000000000) {
    words = `${convertDigitsToWords(Math.floor(num / 10000000))} Crore ${convertDigitsToWords(num % 10000000)}`;
  }
  return words.trim();
}

const convertToIndianWords = (numStr) => {
  const num = parseFloat(numStr);
  if (isNaN(num) || num > 999999999.99) {
    return "Error: Number too large";
  }
  if (num === 0) {
    return "Rupees Zero only";
  }
  const numArr = numStr.split('.');
  let words = '';
  if (numArr[0] !== '0') {
    words += `Rupees ${convertDigitsToWords(parseInt(numArr[0], 10))} `;
  } else {
    words += 'Rupees ';
  }
  if (numArr[1] !== '00') {
    words += `and ${convertDigitsToWords(parseInt(numArr[1], 10))} Paise only`;
  } else {
    words += 'only';
  }
  return words;
};

export { convertToIndianWords };