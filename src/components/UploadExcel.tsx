import { useState, ChangeEvent } from "react";
import { usePapaParse, ParseResult } from "react-papaparse";
import * as XLSX from "xlsx";
import {
  Table, TableHead, TableRow, TableCell, TableBody, Button, Box,
  Pagination, Alert, CircularProgress, Typography, Paper,
  TableContainer, Chip, LinearProgress
} from "@mui/material";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";
import DeleteIcon from "@mui/icons-material/Delete";
import UploadIcon from "@mui/icons-material/Upload";
import {searchContactByEmail, insertAccountRecord, insertContactRecord, insertDealRecord, updateContactRecord} from "../services/Zoho";

type CellValue = string | number | null | undefined;
type RowData = CellValue[];
type TableData = RowData[];

interface FileInfo {
  name: string;
  size: number;
  type: string;
  rows: number;
}

const UploadExcel = () => {
  const [data, setData] = useState<TableData>([]);
  const [pageNumber, setPageNumber] = useState<number>(1);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [fileInfo, setFileInfo] = useState<FileInfo | null>(null);
  const [importing, setImporting] = useState<boolean>(false);
  const [importProgress, setImportProgress] = useState<number>(0);
  const [importStatus, setImportStatus] = useState<string | null>(null);

  const rowsPerPage = 25;
  const pageCount = Math.ceil((data.length > 0 ? data.length - 1 : 0) / rowsPerPage);
  const { readString } = usePapaParse();

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + ' bytes';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
  };

  const handleFileUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const validTypes = [
      'text/csv',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ];

    if (!validTypes.includes(file.type) &&
        !file.name.endsWith('.csv') &&
        !file.name.endsWith('.xlsx')) {
      setError("Invalid file type. Please upload a CSV or XLSX file.");
      return;
    }

    setLoading(true);
    setError(null);
    setData([]);
    setPageNumber(1);

    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const binaryStr = e.target?.result as string;

        if (file.name.endsWith(".csv")) {
          readString(binaryStr, {
            complete: (result: ParseResult<string[]>) => {
              const parsedData = result.data;
              setData(parsedData);
              setFileInfo({
                name: file.name,
                size: file.size,
                type: 'CSV',
                rows: parsedData.length - 1
              });
              setLoading(false);
            },
            error: (error) => {
              setError(`Error parsing CSV: ${error.message}`);
              setLoading(false);
            },
            skipEmptyLines: true,
          });
        } else if (file.name.endsWith(".xlsx")) {
          try {
            const workbook = XLSX.read(binaryStr, { type: "binary" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const parsedData = XLSX.utils.sheet_to_json<RowData>(sheet, { header: 1 });
            setData(parsedData);
            setFileInfo({
              name: file.name,
              size: file.size,
              type: 'XLSX',
              rows: parsedData.length - 1
            });
            setLoading(false);
          } catch (error) {
            setError(`Error parsing XLSX: ${error instanceof Error ? error.message : 'Unknown error'}`);
            setLoading(false);
          }
        }
      } catch (error) {
        setError(`Error reading file: ${error instanceof Error ? error.message : 'Unknown error'}`);
        setLoading(false);
      }
    };

    reader.onerror = () => {
      setError("Error reading file");
      setLoading(false);
    };

    reader.readAsBinaryString(file);

    event.target.value = '';
  };

  const handleImport = async () => {
    if (data.length <= 1) return;

    setImporting(true);
    setImportProgress(0);
    setImportStatus("Starting import...");

    const rowsToProcess = data.slice(1);
    const totalRows = rowsToProcess.length;
    let processed = 0;
    let successful = 0;
    let failures = 0;

    try {
      const headers = data[0];
      const emailIndex = headers.findIndex(h => String(h).includes('Homeowners Email'));
      const fullNameIndex = headers.findIndex(h => String(h).includes('Name'));
      const accountNameIndex = headers.findIndex(h => String(h).includes('Name'));

      const doeIdIndex = headers.findIndex(h => String(h).includes('DOE ID Number'));
        const dealNameIndex = headers.findIndex(h => String(h).includes('Name'));
        const latitudeIndex = headers.findIndex(h => String(h).includes('Latitude'));
        const longitudeIndex = headers.findIndex(h => String(h).includes('Longitude'));
        const phoneNumberIndex = headers.findIndex(h => String(h).includes('Phone number'));
        const municipalityIndex = headers.findIndex(h => String(h).includes('Municipality'));
        const houseNumberIndex = headers.findIndex(h => String(h).includes('House Number and Street Name (INSTALLATION)'));
        const zipCodeIndex = headers.findIndex(h => String(h).includes('Zip Code (INSTALLATION)'));
        const cityIndex = headers.findIndex(h => String(h).includes('City (INSTALLATION)'));
        const constructionYearIndex = headers.findIndex(h => String(h).includes('Construction Year (enter 4 digit year ex. 1950)'));
        const houseAgeIndex = headers.findIndex(h => String(h).includes('Is the single dwelling house 50 years of age or older?'));
        const roofTypeIndex = headers.findIndex(h => String(h).includes('Does the house have a flat or inclined/pitched roof?'));
        const disabilityIndex = headers.findIndex(h => String(h).includes('Is anyone in the household eligible as an Energy Dependent Disability Individual'));
        const medicalEquipmentIndex = headers.findIndex(h => String(h).includes('If your medical equipment is not listed above'));
        const medDeviceUploadIndex = headers.findIndex(h => String(h).includes('Upload of Med Device'));
        const assignedIndex = headers.findIndex(h => String(h).includes('Assigned'));

        const mailingHouseNumberIndex = headers.findIndex(h => String(h).includes('House Number and Street Name (MAILING)'));
        const mailingCityIndex = headers.findIndex(h => String(h).includes('City (MAILING)'));
        const mailingZipCodeIndex = headers.findIndex(h => String(h).includes('Zip Code (MAILING)'));
        const mailingMunicipalityIndex = headers.findIndex(h => String(h).includes('Municipality (MAILING)'));

      if (emailIndex === -1) {
        throw new Error("Email column not found in the data");
      }

      for (const row of rowsToProcess) {
        try {
          const email = row[emailIndex]?.toString() || '';

          let firstName = '';
          let lastName = '';
          if (fullNameIndex !== -1) {
            const fullName = row[fullNameIndex]?.toString() || '';
            const nameParts = fullName.split(' ');
            if (nameParts.length > 2) {
              firstName = nameParts.slice(0, -2).join(' ');
              lastName = nameParts.slice(-2).join(' ');
            } else if (nameParts.length === 2) {
              firstName = nameParts[0];
              lastName = nameParts[1];
            } else {
              firstName = fullName;
            }
          }

          const accountName = accountNameIndex !== -1 ? row[accountNameIndex]?.toString().trim().replace(/\s+/g, ' ') : '';
          const doeId = doeIdIndex !== -1 ? row[doeIdIndex]?.toString() : '';
          const dealName = dealNameIndex !== -1 ? row[dealNameIndex]?.toString() : '';
          const latitude = latitudeIndex !== -1 ? row[latitudeIndex]?.toString() : '';
          const longitude = longitudeIndex !== -1 ? row[longitudeIndex]?.toString() : '';
          const phoneNumber = phoneNumberIndex !== -1 ? row[phoneNumberIndex]?.toString() : '';
          const municipality = municipalityIndex !== -1 ? row[municipalityIndex]?.toString() : '';
          const houseNumber = houseNumberIndex !== -1 ? row[houseNumberIndex]?.toString() : '';
          const zipCode = zipCodeIndex !== -1 ? row[zipCodeIndex]?.toString() : '';
          const city = cityIndex !== -1 ? row[cityIndex]?.toString() : '';
          const constructionYear = constructionYearIndex !== -1 ? row[constructionYearIndex]?.toString() : '';
          const houseAge = houseAgeIndex !== -1 ? row[houseAgeIndex]?.toString() : '';
          const roofType = roofTypeIndex !== -1 ? row[roofTypeIndex]?.toString() : '';
          const disability = disabilityIndex !== -1 ? row[disabilityIndex]?.toString() : '';
          const medicalEquipment = medicalEquipmentIndex !== -1 ? row[medicalEquipmentIndex]?.toString() : '';
          const medDeviceUpload = medDeviceUploadIndex !== -1 ? row[medDeviceUploadIndex]?.toString() : '';
          const assigned = assignedIndex !== -1 && row[assignedIndex] ? new Date((row[assignedIndex] as number - (25567 + 1)) * 86400 * 1000).toISOString().split('T')[0] : '';

          const mailingHouseNumber = disabilityIndex !== -1 ? row[mailingHouseNumberIndex]?.toString() : '';
          const mailingCity = medicalEquipmentIndex !== -1 ? row[mailingCityIndex]?.toString() : '';
          const mailingZipCode = mailingZipCodeIndex !== -1 ? row[mailingZipCodeIndex]?.toString() : '';
          const mailingMunicipality = mailingMunicipalityIndex !== -1 && row[mailingMunicipalityIndex] ? row[mailingMunicipalityIndex].toString() : '';

          if (!email) {
            console.warn("Skipping row with empty email", row);
            processed++;
            failures++;
            continue;
          }

          setImportStatus(`Processing ${email}...`);

          // Step 1: Search for contact by email
          const contactResult = await searchContactByEmail(email);
          console.log(`Contact search result for ${email}:`, contactResult);

            let contactId;
            let accountId;

            // Step 3: If contact not found, insert contact record
            if (!contactResult.found) {
                 // Step 2: Insert account record
                const accountData = { "Account_Name": accountName };
                console.log("insertAccountRecord:", accountData);
                const accountResult = await insertAccountRecord(accountData);
                console.log("accountResult: ", accountResult);

                if (!accountResult.success) {
                    console.error(`Failed to create account for ${email}:`, accountResult.error);
                    throw new Error(`Failed to create account: ${accountResult.error}`);
                }

                accountId = accountResult.data[0].details.id;

                const contactData = {
                    "Email": email,
                    "First_Name": firstName,
                    "Last_Name": lastName,
                    "Mailing Street": mailingHouseNumber,
                    "Mailing_City": mailingCity,
                    "Mailing_Zip": mailingZipCode,
                    "County": mailingMunicipality,
                    "Account_Name": accountId,
                };
                console.log("insertContactRecord:", contactData);

                const newContactResult = await insertContactRecord(contactData);

                if (!newContactResult.success) {
                console.error(`Failed to create contact for ${email}:`, newContactResult.error);
                throw new Error(`Failed to create contact: ${newContactResult.error}`);
                }

                contactId = newContactResult.data[0].details.id;
            } else {
                contactId = contactResult.data[0].id;
                accountId = contactResult.data[0].Account_Name || null;

                if (!accountId) {
                    // Step 2: Insert account record
                    const accountData = { "Account_Name": accountName };
                    const accountResult = await insertAccountRecord(accountData);
                    console.log("insertAccountRecord:", accountData);

                    if (!accountResult.success) {
                        console.error(`Failed to create account for ${email}:`, accountResult.error);
                        throw new Error(`Failed to create account: ${accountResult.error}`);
                    }

                    accountId = accountResult.data[0].details.id;

                    const contactData = {
                        "Account_Name": accountId
                    };

                    await updateContactRecord(contactData);
                }
            }

            // Step 4: Insert deal record
            const dealData = {
                "DOE_ID_Number": doeId,
                "Deal_Name": dealName,
                "Latitude": latitude,
                "Longitude": longitude,
                "Customer_Phone": phoneNumber,
                "Layout": { "id": "4909080000146647839" },
                "Tipo": "Comercial",
                "Customer_State  ": municipality,
                "Customer_Street ": houseNumber,
                "Customer_Postal_Code  ": zipCode,
                "Customer_City": city,
                "Construction_Year": constructionYear,
                "Is_the_single_dwelling_house_50_yrs_or_older": houseAge,
                "the_house_have_a_flat_or_inclined_pitched_roof": roofType,
                "Is_anyone_eligible_as_an_Energy_Dependent": disability,
                "If_your_medical_equipment_is_not_listed_above": medicalEquipment,
                "Upload_of_Med_Device": medDeviceUpload,
                "Assigned": assigned,
                "Contact_Name": contactId,
                "Account_Name": accountId,
                "Stage": "New"
            };
            console.log("insertDealRecord:", dealData);

            const dealResult = await insertDealRecord(dealData);

            if (!dealResult.success) {
                console.error(`Failed to create deal for ${email}:`, dealResult.error);
                throw new Error(`Failed to create deal: ${dealResult.error}`);
            }

            successful++;
        } catch (error) {
          console.error("Error processing row:", error);
          failures++;
        } finally {
          processed++;
          setImportProgress(Math.round((processed / totalRows) * 100));
          setImportStatus(`Processed ${processed} of ${totalRows} rows (${successful} successful, ${failures} failed)`);
        }
      }

      setImportStatus(`Import completed: ${successful} successful, ${failures} failed out of ${totalRows} total records`);
    } catch (error) {
      setImportStatus(`Import failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setError(`Import failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setImporting(false);
    }
  };

  const handlePageChange = (_event: React.ChangeEvent<unknown>, value: number) => {
    setPageNumber(value);
  };

  const resetData = () => {
    setData([]);
    setPageNumber(1);
    setFileInfo(null);
    setError(null);
  };

  const displayRows = data.length > 0 ?
    data.slice(1).slice((pageNumber - 1) * rowsPerPage, pageNumber * rowsPerPage) : [];

  const hasHeaders = data.length > 0;
  const hasData = data.length > 1;

  return (
    <Paper elevation={3} className="p-4" sx={{ maxWidth: '100%', overflowX: 'auto', m: 2, p: 2 }}>
      <Box display="flex" alignItems="center" justifyContent="center" mb={2} gap={2}>
        <input
          type="file"
          accept=".csv,.xlsx"
          onChange={handleFileUpload}
          style={{ display: "none" }}
          id="file-upload"
          disabled={loading || importing}
        />
        <label htmlFor="file-upload">
          <Button
            variant="contained"
            component="span"
            color="primary"
            startIcon={<CloudUploadIcon />}
            disabled={loading || importing}
          >
            Upload File
          </Button>
        </label>

        {hasData && (
          <>
            <Button
              variant="contained"
              color="success"
              startIcon={<UploadIcon />}
              onClick={handleImport}
              disabled={loading || importing}
            >
              Import
            </Button>
            <Button
              variant="outlined"
              color="error"
              startIcon={<DeleteIcon />}
              onClick={resetData}
              disabled={loading || importing}
            >
              Clear
            </Button>
          </>
        )}
      </Box>

      {error && (
        <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>
      )}

      {importing && (
        <Box sx={{ width: '100%', mb: 2 }}>
          <LinearProgress variant="determinate" value={importProgress} color="success" />
          <Typography variant="body2" color="text.secondary" align="center" mt={1}>
            {importStatus}
          </Typography>
        </Box>
      )}

      {!importing && importStatus && (
        <Alert severity="success" sx={{ mb: 2 }}>{importStatus}</Alert>
      )}

      {loading && (
        <Box display="flex" justifyContent="center" my={4}>
          <CircularProgress />
        </Box>
      )}

      {fileInfo && (
        <Box mb={2} mt={2}>
          <Typography variant="subtitle1" gutterBottom>
            Loaded file: {fileInfo.name}
          </Typography>
          <Box display="flex" gap={1} flexWrap="wrap">
            <Chip label={`Type: ${fileInfo.type}`} color="primary" size="small" />
            <Chip label={`Size: ${formatFileSize(fileInfo.size)}`} color="primary" size="small" />
            <Chip label={`Rows: ${fileInfo.rows}`} color="primary" size="small" />
          </Box>
        </Box>
      )}

      {!loading && !hasData && !error && (
        <Box textAlign="center" my={6}>
          <Typography color="textSecondary">
            No data to display. Please upload a CSV or XLSX file.
          </Typography>
        </Box>
      )}

      {hasData && (
        <>
          <TableContainer component={Paper} sx={{ maxHeight: "70vh" }}>
            <Table size="small">
              <TableHead>
                <TableRow>
                  {data[0].map((header, index) => (
                    <TableCell
                      key={index}
                      sx={{ fontWeight: 'bold', bgcolor: 'primary.light', color: 'white' }}
                    >
                      {String(header || `Column ${index + 1}`)}
                    </TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {displayRows.map((row, rowIndex) => (
                  <TableRow key={rowIndex} hover>
                    {row.map((cell, cellIndex) => (
                      <TableCell key={cellIndex}>
                        {String(cell ?? '')}
                      </TableCell>
                    ))}
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>

          {pageCount > 1 && (
            <Box display="flex" justifyContent="center" alignItems="center" mt={3}>
              <Pagination
                count={pageCount}
                page={pageNumber}
                onChange={handlePageChange}
                color="primary"
                shape="rounded"
                size="large"
              />
            </Box>
          )}
        </>
      )}
    </Paper>
  );
};

export default UploadExcel;
