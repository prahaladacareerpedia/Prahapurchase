import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

function App() {
  const [data, setData] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      setData(jsonData);
      console.log(jsonData);
    };
    reader.readAsBinaryString(file);
  };

  const generateTallyXML = () => {
    if (!data.length) return;

    const createXml = (data) => {
      const xmlDoc = document.implementation.createDocument('', '', null);
      const envelope = xmlDoc.createElement('ENVELOPE');

      const header = xmlDoc.createElement('HEADER');
      const tallyRequest = xmlDoc.createElement('TALLYREQUEST');
      tallyRequest.textContent = 'Import Data';
      header.appendChild(tallyRequest);
      envelope.appendChild(header);

      const body = xmlDoc.createElement('BODY');
      const importData = xmlDoc.createElement('IMPORTDATA');
      const requestDesc = xmlDoc.createElement('REQUESTDESC');
      const reportName = xmlDoc.createElement('REPORTNAME');
      reportName.textContent = 'Vouchers';
      requestDesc.appendChild(reportName);

      const staticVars = xmlDoc.createElement('STATICVARIABLES');
      const svcCompany = xmlDoc.createElement('SVCURRENTCOMPANY');
      svcCompany.textContent = 'Your Company Name';
      staticVars.appendChild(svcCompany);
      requestDesc.appendChild(staticVars);
      importData.appendChild(requestDesc);

      const requestData = xmlDoc.createElement('REQUESTDATA');

      const groupedData = groupBy(data, 'Invoice number'); 

      for (const [invoiceNumber, group] of Object.entries(groupedData)) {
        const tallyMessage = xmlDoc.createElement('TALLYMESSAGE');
        const voucher = xmlDoc.createElement('VOUCHER');
        voucher.setAttribute('VCHTYPE', 'Purchase');
        voucher.setAttribute('ACTION', 'Create');
        voucher.setAttribute('OBJVIEW', 'Accounting Voucher View');

        const firstEntry = group[0];
        const formattedDate = formatDate(firstEntry['Invoice Date']); 

        createElementWithText(voucher, 'DATE', formattedDate);
        createElementWithText(voucher, 'REFERENCEDATE', formattedDate);
        createElementWithText(voucher, 'VCHSTATUSDATE', formattedDate);
        createElementWithText(voucher, 'GUID', 'your_guid_here');
        createElementWithText(voucher, 'VCHKEY', 'your_vchkey_here');
        createElementWithText(voucher, 'GSTREGISTRATIONTYPE', 'Regular');
        createElementWithText(voucher, 'STATENAME', firstEntry['Place of supply']);
        createElementWithText(voucher, 'COUNTRYOFRESIDENCE', 'India');
        createElementWithText(voucher, 'PARTYGSTIN', firstEntry['GSTIN of supplier']);
        createElementWithText(voucher, 'PARTYNAME', firstEntry['Trade/Legal name']);
        createElementWithText(voucher, 'CMPGSTIN', '37AAIFR7081C1Z3');
        createElementWithText(voucher, 'VOUCHERTYPENAME', 'Purchase');
        createElementWithText(voucher, 'PARTYLEDGERNAME', firstEntry['Trade/Legal name']);
        createElementWithText(voucher, 'VOUCHERNUMBER', invoiceNumber);
        createElementWithText(voucher, 'SUPPLIERINVOICENO', invoiceNumber);
        createElementWithText(voucher, 'REFERENCE', invoiceNumber);

        const partyLedgerEntry = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
        createElementWithText(partyLedgerEntry, 'LEDGERNAME', firstEntry['Trade/Legal name']);
        createElementWithText(partyLedgerEntry, 'ISDEEMEDPOSITIVE', 'No');
        const partyAmount = Math.round(
          group.reduce((sum, row) => sum + row['Taxable Value (₹)'] + row['Central Tax(₹)'] + row['State/UT Tax(₹)'], 0) * 100
        ) / 100;
        createElementWithText(partyLedgerEntry, 'AMOUNT', partyAmount);

        const billAlloc = xmlDoc.createElement('BILLALLOCATIONS.LIST');
        createElementWithText(billAlloc, 'NAME', invoiceNumber);
        createElementWithText(billAlloc, 'BILLTYPE', 'New Ref');
        createElementWithText(billAlloc, 'AMOUNT', partyAmount);
        partyLedgerEntry.appendChild(billAlloc);

        voucher.appendChild(partyLedgerEntry);

        group.forEach(row => {
          const ledgerEntry = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
          const rate = row['Rate(%)'];
          const ledgerName = `${rate}`;  // Use exactly as in Excel
          createElementWithText(ledgerEntry, 'LEDGERNAME', ledgerName);
          createElementWithText(ledgerEntry, 'ISDEEMEDPOSITIVE', 'Yes');
          createElementWithText(ledgerEntry, 'AMOUNT', -row['Taxable Value (₹)']);
          voucher.appendChild(ledgerEntry);

          const centralTaxEntry = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
          createElementWithText(centralTaxEntry, 'LEDGERNAME', 'Central Tax');
          createElementWithText(centralTaxEntry, 'ISDEEMEDPOSITIVE', 'Yes');
          createElementWithText(centralTaxEntry, 'AMOUNT', -row['Central Tax(₹)']);
          voucher.appendChild(centralTaxEntry);

          const stateTaxEntry = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
          createElementWithText(stateTaxEntry, 'LEDGERNAME', 'State/UT Tax');
          createElementWithText(stateTaxEntry, 'ISDEEMEDPOSITIVE', 'Yes');
          createElementWithText(stateTaxEntry, 'AMOUNT', -row['State/UT Tax(₹)']);
          voucher.appendChild(stateTaxEntry);
        });

        tallyMessage.appendChild(voucher);
        requestData.appendChild(tallyMessage);
      }

      importData.appendChild(requestData);
      body.appendChild(importData);
      envelope.appendChild(body);
      xmlDoc.appendChild(envelope);

      const serializer = new XMLSerializer();
      return serializer.serializeToString(xmlDoc);
    };

    const xmlContent = createXml(data);
    const blob = new Blob([xmlContent], { type: 'application/xml' });
    saveAs(blob, 'TallyData.xml');
  };

  const createElementWithText = (parent, tagName, text) => {
    const element = document.createElement(tagName);
    element.textContent = text;
    parent.appendChild(element);
  };

  const groupBy = (array, key) => {
    return array.reduce((result, currentValue) => {
      (result[currentValue[key]] = result[currentValue[key]] || []).push(currentValue);
      return result;
    }, {});
  };

  const formatDate = (dateValue) => {
    if (typeof dateValue === 'number') {
      const parsedDate = new Date((dateValue - 25569) * 86400 * 1000);
      const year = parsedDate.getFullYear();
      const month = (`0${parsedDate.getMonth() + 1}`).slice(-2);
      const day = (`0${parsedDate.getDate()}`).slice(-2);
      return `${year}${month}${day}`;
    } else if (typeof dateValue === 'string') {
      const [day, month, year] = dateValue.split('/');
      return `${year}${month}${day}`;
    } else {
      throw new Error(`Unexpected date format: ${dateValue}`);
    }
  };

  return (
    <div className="App">
      <h1>Excel to Tally XML Converter - Purchase</h1>
      <input type="file" onChange={handleFileUpload} />
      <button onClick={generateTallyXML}>Generate Tally XML</button>
    </div>
  );
}

export default App;
