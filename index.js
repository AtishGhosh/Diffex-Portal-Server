import express from 'express';
import generateInvoice from './invoicing/generate-invoice.js';
import getLatestData from './invoicing/latest-data.js';

const app = express();
const port = 3000;

app.use(express.json());

// Set up CORS headers
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

app.post('/generate-invoice', generateInvoice);

app.post('/latest-data', getLatestData);

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
