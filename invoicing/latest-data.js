import fs from 'fs';
import path from 'path';

const filePath = path.join(__dirname, '../storage/data/invoices.json');

function getLatestData(req, res) {
  fs.readFile(filePath, 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Unable to read file');
      return;
    }

    try {
      const invoices = JSON.parse(data);
      const latestData = invoices[invoices.length - 1];
      res.json(latestData);
    } catch (error) {
      console.error(error);
      res.status(500).send('Unable to parse file contents');
    }
  });
}

export default getLatestData;
