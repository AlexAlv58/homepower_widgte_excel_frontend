import React, { useState, ChangeEvent, useMemo } from "react";
import { usePapaParse, ParseResult } from "react-papaparse";
import * as XLSX from "xlsx";
import {
  Table, TableHead, TableRow, TableCell, TableBody, Button, Box,
  Pagination, Alert, CircularProgress, Typography, Paper,
  TableContainer, Chip, LinearProgress, Dialog, DialogTitle,
  DialogContent, DialogActions, List, ListItem, ListItemIcon,
  ListItemText, Divider, Tooltip, useTheme
} from "@mui/material";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";
import DeleteIcon from "@mui/icons-material/Delete";
import UploadIcon from "@mui/icons-material/Upload";
import DownloadIcon from "@mui/icons-material/Download";
import InfoIcon from "@mui/icons-material/Info";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import AccountCircleIcon from "@mui/icons-material/AccountCircle";
import BusinessIcon from "@mui/icons-material/Business";
import EmailIcon from "@mui/icons-material/Email";
import WarningIcon from "@mui/icons-material/Warning";
import FilterListIcon from '@mui/icons-material/FilterList';
import FilterListOffIcon from '@mui/icons-material/FilterListOff';
import PlaylistRemoveIcon from '@mui/icons-material/PlaylistRemove';
import ArrowForwardIcon from '@mui/icons-material/ArrowForward';
import {searchContactByEmail, insertAccountRecord, insertContactRecord,insertSubmoduleRecord ,insertDealRecord, updateContactRecord} from "../services/Zoho";

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
  const [validationErrors, setValidationErrors] = useState<Array<{row: number, email: string, message: string}>>([]);
  const [hasValidationErrors, setHasValidationErrors] = useState<boolean>(false);
  const [infoDialogOpen, setInfoDialogOpen] = useState<boolean>(false);
  const [showOnlyErrors, setShowOnlyErrors] = useState<boolean>(false);
  const [importErrors, setImportErrors] = useState<Array<{rowNumber: number, email: string, error: string}>>([]);

  const theme = useTheme();
  const rowsPerPage = 25;
  const { readString } = usePapaParse();

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + ' bytes';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
  };

  // Función para validar formato de email
  const isValidEmail = (email: string): boolean => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
  };

  // Función para validar todos los registros
  const validateData = (data: TableData): Array<{row: number, email: string, message: string}> => {
    const errors: Array<{row: number, email: string, message: string}> = [];
    
    if (data.length <= 1) return errors;

    const headers = data[0];
    const emailIndex = headers.findIndex(h => String(h).toLowerCase().includes('homeowner') && String(h).toLowerCase().includes('email'));
    
    if (emailIndex === -1) {
      errors.push({
        row: 0,
        email: '',
        message: 'No se encontró la columna de email en los datos'
      });
      return errors;
    }

    const rowsToProcess = data.slice(1);
    
    rowsToProcess.forEach((row, index) => {
      const rowNumber = index + 2; // +2 porque index empieza en 0 y la primera fila son headers
      const email = row[emailIndex]?.toString() || '';
      
      if (!email || email.trim() === '') {
        errors.push({
          row: rowNumber,
          email: '',
          message: 'El email es obligatorio y no puede estar vacío'
        });
      } else if (!isValidEmail(email)) {
        errors.push({
          row: rowNumber,
          email: email,
          message: 'Formato de email inválido'
        });
      }
    });

    return errors;
  };

  // Función para normalizar los datos y manejar celdas vacías
  const normalizeData = (data: TableData): TableData => {
    if (data.length === 0) return data;

    const headers = data[0];
    const maxColumns = Math.max(...data.map(row => row.length));
    
    // Asegurar que todos los headers tengan el mismo número de columnas
    const normalizedHeaders = [...headers];
    while (normalizedHeaders.length < maxColumns) {
      normalizedHeaders.push(`Columna ${normalizedHeaders.length + 1}`);
    }

    // Normalizar todas las filas de datos
    const normalizedRows = data.slice(1).map(row => {
      const normalizedRow = [...row];
      while (normalizedRow.length < maxColumns) {
        normalizedRow.push(null);
      }
      return normalizedRow;
    });

    return [normalizedHeaders, ...normalizedRows];
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
      setError("Tipo de archivo inválido. Por favor, suba un archivo CSV o XLSX.");
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
              const normalizedData = normalizeData(parsedData);
              setData(normalizedData);
              setFileInfo({
                name: file.name,
                size: file.size,
                type: 'CSV',
                rows: normalizedData.length - 1
              });
              
              // Validar datos después de cargar
              const errors = validateData(normalizedData);
              setValidationErrors(errors);
              setHasValidationErrors(errors.length > 0);
              
              setLoading(false);
            },
            error: (error) => {
              setError(`Error al procesar CSV: ${error.message}`);
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
            const normalizedData = normalizeData(parsedData);
            setData(normalizedData);
            setFileInfo({
              name: file.name,
              size: file.size,
              type: 'XLSX',
              rows: normalizedData.length - 1
            });
            
            // Validar datos después de cargar
            const errors = validateData(normalizedData);
            setValidationErrors(errors);
            setHasValidationErrors(errors.length > 0);
            
            setLoading(false);
          } catch (error) {
            setError(`Error al procesar XLSX: ${error instanceof Error ? error.message : 'Error desconocido'}`);
            setLoading(false);
          }
        }
      } catch (error) {
        setError(`Error al leer archivo: ${error instanceof Error ? error.message : 'Error desconocido'}`);
        setLoading(false);
      }
    };

    reader.onerror = () => {
      setError("Error al leer archivo");
      setLoading(false);
    };

    reader.readAsBinaryString(file);

    event.target.value = '';
  };

  const handleRemoveErrors = () => {
    if (!hasValidationErrors) return;

    const errorRowNumbers = new Set(validationErrors.map(e => e.row));
    const newData = [
      data[0],
      ...data.slice(1).filter((_row, index) => !errorRowNumbers.has(index + 2))
    ];

    setData(newData);
    setValidationErrors([]);
    setHasValidationErrors(false);
    if (showOnlyErrors) {
      setShowOnlyErrors(false);
    }
    setFileInfo(prev => prev ? { ...prev, rows: newData.length - 1 } : null);
    setPageNumber(1);
  };

  const downloadErrorFile = () => {
    if (!hasValidationErrors) return;

    const errorRows = validationErrors.map(e => data[e.row - 1]);
    const dataToExport = [data[0], ...errorRows];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(dataToExport);

    const columnWidths = data[0].map(header => ({ wch: Math.max(String(header).length, 15) }));
    worksheet['!cols'] = columnWidths;

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filas con Errores');
    XLSX.writeFile(workbook, 'registros_con_errores.xlsx');
  };

  const handleProcessErrors = () => {
    downloadErrorFile();
    handleRemoveErrors();
  };

  const handleImport = async () => {
    if (data.length <= 1) return;

    // Verificar si hay errores de validación
    if (hasValidationErrors) {
      setError("No se puede proceder con la importación. Por favor, corrija los errores de validación primero.");
      return;
    }

    setImporting(true);
    setImportProgress(0);
    setImportStatus("Iniciando importación...");
    setImportErrors([]); // Limpiar errores de importaciones anteriores

    const rowsToProcess = data.slice(1);
    const totalRows = rowsToProcess.length;
    let processed = 0;
    let successful = 0;
    let failures = 0;
    const newImportErrors: Array<{rowNumber: number, email: string, error: string}> = [];

    try {
      const headers = data[0];
      
      // Nuevas columnas principales
      const dateAssignedIndex = headers.findIndex(h => String(h).toLowerCase().includes('date assigned'));
      const numberIndex = headers.findIndex(h => String(h).toLowerCase().includes('number') && !String(h).toLowerCase().includes('phone'));
      const nameIndex = headers.findIndex(h => String(h).toLowerCase().includes('name'));
      const latitudeIndex = headers.findIndex(h => String(h).toLowerCase().includes('latitude'));
      const longitudeIndex = headers.findIndex(h => String(h).toLowerCase().includes('longitude'));
      const phoneNumberIndex = headers.findIndex(h => String(h).toLowerCase().includes('phone number') && !String(h).toLowerCase().includes('alternate'));
      const alternatePhoneIndex = headers.findIndex(h => String(h).toLowerCase().includes('alternate phone number'));
      const emailIndex = headers.findIndex(h => String(h).toLowerCase().includes('homeowner') && String(h).toLowerCase().includes('email'));
      const municipalityIndex = headers.findIndex(h => String(h).toLowerCase().includes('municipality') && !String(h).toLowerCase().includes('mailing'));
      
      // Dirección de instalación (sin sufijo)
      const houseNumberIndex = headers.findIndex(h => String(h).toLowerCase().includes('house number and street name') && !String(h).toLowerCase().includes('mailing'));
      const cityIndex = headers.findIndex(h => String(h).toLowerCase().includes('city') && !String(h).toLowerCase().includes('mailing'));
      const zipCodeIndex = headers.findIndex(h => String(h).toLowerCase().includes('zip code') && !String(h).toLowerCase().includes('mailing'));
      
      // Dirección de correo
      const mailingHouseNumberIndex = headers.findIndex(h => String(h).toLowerCase().includes('house number and street name (mailing)'));
      const mailingCityIndex = headers.findIndex(h => String(h).toLowerCase().includes('city (mailing)'));
      const mailingMunicipalityIndex = headers.findIndex(h => String(h).toLowerCase().includes('municipality (mailing)'));
      const mailingZipCodeIndex = headers.findIndex(h => String(h).toLowerCase().includes('zip code (mailing)'));
      
      // Información de la casa
      const constructionYearIndex = headers.findIndex(h => String(h).toLowerCase().includes('construction year'));
      const houseAgeIndex = headers.findIndex(h => String(h).toLowerCase().includes('50 years of age or older'));
      const roofTypeIndex = headers.findIndex(h => String(h).toLowerCase().includes('flat or inclined/pitched roof'));
      const housecementConcreteIndex = headers.findIndex(h => String(h).toLowerCase().includes('cement/concrete or metal/zinc'));
      const geographicEligibilityIndex = headers.findIndex(h => String(h).toLowerCase().includes('geographic eligibility (last mile community'));
      
      // Información de discapacidad
      const disabilityIndex = headers.findIndex(h => String(h).toLowerCase().includes('energy dependent disability individual'));
      const medicalEquipmentIndex = headers.findIndex(h => String(h).toLowerCase().includes('medical equipment is not listed above'));
      const energyEligibilityIndex = headers.findIndex(h => String(h).toLowerCase().includes('energy dependent disability eligibility'));
      const noConsentPicturesIndex = headers.findIndex(h => String(h).toLowerCase().includes('did not consent to pictures'));

      // Equipos médicos
      const airConditionerIndex = headers.findIndex(h => String(h).toLowerCase().includes('air conditioner (a/c) for temperature control'));
      const airPurifierIndex = headers.findIndex(h => String(h).toLowerCase().includes('air purifier'));
      const airMattressIndex = headers.findIndex(h => String(h).toLowerCase().includes('air mattress for bed sores'));
      const asthmaTherapyMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('asthma therapy machine or nebulizer'));
      const atHomeDialysisMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('at home dialysis machine'));
      const bipapMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('bilevel positive airway pressure (bipap) machine'));
      const sleepApneaMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('cpap, bpap, apap or any other sleep apnea machine'));
      const dehumidifierIndex = headers.findIndex(h => String(h).toLowerCase().includes('dehumidifier'));
      const dialysisMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('dialysis machine'));
      const electricCraneIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric crane'));
      const physicalTherapyMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric machine for physical therapy'));
      const powerLiftReclinerIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric power lift recliner'));
      const vitalSignsMonitorIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric vital signs monitor'));
      const electricBedIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric bed equipment in the last 13 months'));
      const electricScooterIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric scooter'));
      const electricWheelchairIndex = headers.findIndex(h => String(h).toLowerCase().includes('electric wheelchair'));
      const enteralFeedingPumpIndex = headers.findIndex(h => String(h).toLowerCase().includes('enteral feeding tube pump machine / naso feeding machine'));
      const enteralFeedingMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('enteral feeding machine'));
      const externalDefibrillatorIndex = headers.findIndex(h => String(h).toLowerCase().includes('external defibrillator'));
      const fftMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('fft electric machine'));
      const fanIndex = headers.findIndex(h => String(h).toLowerCase().includes('fan for temperature control'));
      const hearingAidPodsIndex = headers.findIndex(h => String(h).toLowerCase().includes('hearing aid pods rechargeable'));
      const humidifierIndex = headers.findIndex(h => String(h).toLowerCase().includes('humidifier'));
      const implantedCardiacDeviceIndex = headers.findIndex(h => String(h).toLowerCase().includes('implanted cardiac devices that include left ventricular assistive device'));
      const mechanicalVentilatorIndex = headers.findIndex(h => String(h).toLowerCase().includes('mechanical ventilator'));
      const refrigeratedMedicationsIndex = headers.findIndex(h => String(h).toLowerCase().includes('medications that require refrigeration'));
      const oxygenConcentratorIndex = headers.findIndex(h => String(h).toLowerCase().includes('oxegen concentrator equipment in the past 36 months'));
      const neurostimulatorImplantIndex = headers.findIndex(h => String(h).toLowerCase().includes('rechargeable electrical neurostimulator implant'));
      const spinalCordStimulatorIndex = headers.findIndex(h => String(h).toLowerCase().includes('rechargeable spinal cord simulator (scs)'));
      const rvadIndex = headers.findIndex(h => String(h).toLowerCase().includes('right ventricular assistive device (rvad)'));
      const suctionPumpMachineIndex = headers.findIndex(h => String(h).toLowerCase().includes('suction pump machine'));
      const suctionPumpIndex = headers.findIndex(h => String(h).toLowerCase().includes('suction pump'));
      const deafCommunicationIndex = headers.findIndex(h => String(h).toLowerCase().includes('telephone communication for deaf/hard of hearing'));
      const artificialHeartIndex = headers.findIndex(h => String(h).toLowerCase().includes('total artifical heart (tah) in the past 5 years'));
      const vaporizerIndex = headers.findIndex(h => String(h).toLowerCase().includes('vaporizer'));
      const bivadIndex = headers.findIndex(h => String(h).toLowerCase().includes('bi-ventricular assistive device (bivad)'));
      const ivInfusionPumpIndex = headers.findIndex(h => String(h).toLowerCase().includes('intravenous (iv) infusion pump'));

      if (emailIndex === -1) {
        throw new Error("No se encontró la columna de email en los datos");
      }

      for (let i = 0; i < rowsToProcess.length; i++) {
        const row = rowsToProcess[i];
        const rowNumber = i + 2;
        let email = '';

        try {
          // Extraer datos básicos
          let dateAssigned = '';
          const dateValue = row[dateAssignedIndex];

          if (dateAssignedIndex !== -1 && dateValue) {
              let finalDate;
              // Si es un número (formato interno de Excel), lo convertimos
              if (typeof dateValue === 'number' && dateValue > 1) {
                  const d = XLSX.SSF.parse_date_code(dateValue);
                  finalDate = new Date(Date.UTC(d.y, d.m - 1, d.d));
              // Si es texto, intentamos interpretarlo como fecha, ajustando la zona horaria
              } else {
                  const tempDate = new Date(dateValue.toString());
                  finalDate = new Date(tempDate.getTime() + (tempDate.getTimezoneOffset() * 60000));
              }

              if (finalDate instanceof Date && !isNaN(finalDate.getTime())) {
                  dateAssigned = finalDate.toISOString().split("T")[0];
              } else {
                  throw new Error(`Formato de fecha inválido en la columna "Date Assigned"`);
              }
          }
          
          const number = numberIndex !== -1 ? row[numberIndex]?.toString() : '';
          email = row[emailIndex]?.toString() || '';

          if (!email) {
            throw new Error("El email es obligatorio y no puede estar vacío.");
          }

          let firstName = '';
          let lastName = '';
          if (nameIndex !== -1) {
            const fullName = row[nameIndex]?.toString() || '';
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

          // Información de ubicación
          const latitude = latitudeIndex !== -1 ? row[latitudeIndex]?.toString() : '';
          const longitude = longitudeIndex !== -1 ? row[longitudeIndex]?.toString() : '';
          const phoneNumber = phoneNumberIndex !== -1 ? row[phoneNumberIndex]?.toString() : '';
          const alternatePhone = alternatePhoneIndex !== -1 ? row[alternatePhoneIndex]?.toString() : '';
          const municipality = municipalityIndex !== -1 ? row[municipalityIndex]?.toString() : '';
          
          // Dirección de instalación
          const houseNumber = houseNumberIndex !== -1 ? row[houseNumberIndex]?.toString() : '';
          const city = cityIndex !== -1 ? row[cityIndex]?.toString() : '';
          const zipCode = zipCodeIndex !== -1 ? row[zipCodeIndex]?.toString() : '';
          
          // Dirección de correo
          const mailingHouseNumber = mailingHouseNumberIndex !== -1 ? row[mailingHouseNumberIndex]?.toString() : '';
          const mailingCity = mailingCityIndex !== -1 ? row[mailingCityIndex]?.toString() : '';
          const mailingMunicipality = mailingMunicipalityIndex !== -1 ? row[mailingMunicipalityIndex]?.toString() : '';
          const mailingZipCode = mailingZipCodeIndex !== -1 ? row[mailingZipCodeIndex]?.toString() : '';
          
          // Información de la casa
          const constructionYear = constructionYearIndex !== -1 ? row[constructionYearIndex]?.toString() : '';
          const houseAge = houseAgeIndex !== -1 ? row[houseAgeIndex]?.toString() : '';
          const roofType = roofTypeIndex !== -1 ? row[roofTypeIndex]?.toString() : '';
          const housecementConcrete = housecementConcreteIndex !== -1 ? row[housecementConcreteIndex]?.toString() : '';
          const geographicEligibility = geographicEligibilityIndex !== -1 ? row[geographicEligibilityIndex]?.toString() : '';
          
          // Información de discapacidad
          const disability = disabilityIndex !== -1 ? row[disabilityIndex]?.toString() : '';
          const medicalEquipment = medicalEquipmentIndex !== -1 ? row[medicalEquipmentIndex]?.toString() : '';
          const energyEligibility = energyEligibilityIndex !== -1 ? row[energyEligibilityIndex]?.toString() : '';
          const noConsentPictures = noConsentPicturesIndex !== -1 ? row[noConsentPicturesIndex]?.toString() : '';

          // Equipos médicos
          const airConditioner = airConditionerIndex !== -1 ? row[airConditionerIndex]?.toString() : '';
          const airPurifier = airPurifierIndex !== -1 ? row[airPurifierIndex]?.toString() : '';
          const airMattress = airMattressIndex !== -1 ? row[airMattressIndex]?.toString() : '';
          const asthmaTherapyMachine = asthmaTherapyMachineIndex !== -1 ? row[asthmaTherapyMachineIndex]?.toString() : '';
          const atHomeDialysisMachine = atHomeDialysisMachineIndex !== -1 ? row[atHomeDialysisMachineIndex]?.toString() : '';
          const bipapMachine = bipapMachineIndex !== -1 ? row[bipapMachineIndex]?.toString() : '';
          const sleepApneaMachine = sleepApneaMachineIndex !== -1 ? row[sleepApneaMachineIndex]?.toString() : '';
          const dehumidifier = dehumidifierIndex !== -1 ? row[dehumidifierIndex]?.toString() : '';
          const dialysisMachine = dialysisMachineIndex !== -1 ? row[dialysisMachineIndex]?.toString() : '';
          const electricCrane = electricCraneIndex !== -1 ? row[electricCraneIndex]?.toString() : '';
          const physicalTherapyMachine = physicalTherapyMachineIndex !== -1 ? row[physicalTherapyMachineIndex]?.toString() : '';
          const powerLiftRecliner = powerLiftReclinerIndex !== -1 ? row[powerLiftReclinerIndex]?.toString() : '';
          const vitalSignsMonitor = vitalSignsMonitorIndex !== -1 ? row[vitalSignsMonitorIndex]?.toString() : '';
          const electricBed = electricBedIndex !== -1 ? row[electricBedIndex]?.toString() : '';
          const electricScooter = electricScooterIndex !== -1 ? row[electricScooterIndex]?.toString() : '';
          const electricWheelchair = electricWheelchairIndex !== -1 ? row[electricWheelchairIndex]?.toString() : '';
          const enteralFeedingPump = enteralFeedingPumpIndex !== -1 ? row[enteralFeedingPumpIndex]?.toString() : '';
          const enteralFeedingMachine = enteralFeedingMachineIndex !== -1 ? row[enteralFeedingMachineIndex]?.toString() : '';
          const externalDefibrillator = externalDefibrillatorIndex !== -1 ? row[externalDefibrillatorIndex]?.toString() : '';
          const fftMachine = fftMachineIndex !== -1 ? row[fftMachineIndex]?.toString() : '';
          const fan = fanIndex !== -1 ? row[fanIndex]?.toString() : '';
          const hearingAidPods = hearingAidPodsIndex !== -1 ? row[hearingAidPodsIndex]?.toString() : '';
          const humidifier = humidifierIndex !== -1 ? row[humidifierIndex]?.toString() : '';
          const implantedCardiacDevice = implantedCardiacDeviceIndex !== -1 ? row[implantedCardiacDeviceIndex]?.toString() : '';
          const mechanicalVentilator = mechanicalVentilatorIndex !== -1 ? row[mechanicalVentilatorIndex]?.toString() : '';
          const refrigeratedMedications = refrigeratedMedicationsIndex !== -1 ? row[refrigeratedMedicationsIndex]?.toString() : '';
          const oxygenConcentrator = oxygenConcentratorIndex !== -1 ? row[oxygenConcentratorIndex]?.toString() : '';
          const neurostimulatorImplant = neurostimulatorImplantIndex !== -1 ? row[neurostimulatorImplantIndex]?.toString() : '';
          const spinalCordStimulator = spinalCordStimulatorIndex !== -1 ? row[spinalCordStimulatorIndex]?.toString() : '';
          const rvad = rvadIndex !== -1 ? row[rvadIndex]?.toString() : '';
          const suctionPumpMachine = suctionPumpMachineIndex !== -1 ? row[suctionPumpMachineIndex]?.toString() : '';
          const suctionPump = suctionPumpIndex !== -1 ? row[suctionPumpIndex]?.toString() : '';
          const deafCommunication = deafCommunicationIndex !== -1 ? row[deafCommunicationIndex]?.toString() : '';
          const artificialHeart = artificialHeartIndex !== -1 ? row[artificialHeartIndex]?.toString() : '';
          const vaporizer = vaporizerIndex !== -1 ? row[vaporizerIndex]?.toString() : '';
          const bivad = bivadIndex !== -1 ? row[bivadIndex]?.toString() : '';
          const ivInfusionPump = ivInfusionPumpIndex !== -1 ? row[ivInfusionPumpIndex]?.toString() : '';

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
                const accountData = { "Account_Name": firstName + " " + lastName };
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
                    const accountData = { "Account_Name": firstName + " " + lastName };
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

  
            // Función para convertir "true" o "false" a boolean
            // const toBoolean = (value) => value?.toLowerCase() === 'true';
            const toBoolean = (value) => String(value).toLowerCase() === 'true';

            console.log("aire", row[airConditionerIndex]);

            const submModuleData = {
              "Name": firstName + " " + lastName,
              "Air_Conditioner_A_C_for_temperature_control": toBoolean(row[airConditionerIndex]),
              "Air_mattress_for_Bed_Sores_or_Alternating_Air_Pres": toBoolean(row[airMattressIndex]),
              "Air_Purifier": toBoolean(row[airPurifierIndex]),
              "Asthma_therapy_machine_or_Nebulizer": toBoolean(row[asthmaTherapyMachineIndex]),
              "At_home_dialysis_machine": toBoolean(row[atHomeDialysisMachineIndex]),
              "bi_ventricular_assistive_device_BIVAD": toBoolean(row[bivadIndex]),
              "Bilevel_positive_airway_pressure_BiPAP_machine": toBoolean(row[bipapMachineIndex]),
              "CPAP_BPAP_APAP_or_any_other_Sleep_Apnea_Machine": toBoolean(row[sleepApneaMachineIndex]),
              "Dehumidifier": toBoolean(row[dehumidifierIndex]),
              "Dialysis_Machine": toBoolean(row[dialysisMachineIndex]),
              "Electric_bed_equipment_in_the_last_13_months": toBoolean(row[electricBedIndex]),
              "Electric_Crane": toBoolean(row[electricCraneIndex]),
              "Electric_Machine_for_Physical_Therapy": toBoolean(row[physicalTherapyMachineIndex]),
              "Electric_Power_Lift_Recliner": toBoolean(row[powerLiftReclinerIndex]),
              "Electric_scooter": toBoolean(row[electricScooterIndex]),
              "Electric_Vital_Signs_Monitor": toBoolean(row[vitalSignsMonitorIndex]),
              "Electric_wheelchair": toBoolean(row[electricWheelchairIndex]),
              "Energy_Dependent_Disability_Eligibility": toBoolean(row[energyEligibilityIndex]),
              "Energy_Dependent_disability_observed_but_homeowner": toBoolean(row[noConsentPicturesIndex]),
              "Enteral_feeding_machine": toBoolean(row[enteralFeedingMachineIndex]),
              "Enteral_Feeding_Tube_Pump_Machine_Naso_Feeding_M": toBoolean(row[enteralFeedingPumpIndex]),
              "External_Defibrillator": toBoolean(row[externalDefibrillatorIndex]),
              "Fan_for_temperature_control": toBoolean(row[fanIndex]),
              "FFT_Electric_Machine": toBoolean(row[fftMachineIndex]),
              "Hearing_Aid_Pods_Rechargeable": toBoolean(row[hearingAidPodsIndex]),
              "Humidifier": toBoolean(row[humidifierIndex]),
              "Implanted_cardiac_devices_that_include_left_ventri": toBoolean(row[implantedCardiacDeviceIndex]),
              "intravenous_IV_infusion_pump": toBoolean(row[ivInfusionPumpIndex]),
              "Mechanical_Ventilator": toBoolean(row[mechanicalVentilatorIndex]),
              "Medications_that_require_refrigeration": toBoolean(row[refrigeratedMedicationsIndex]),
              "Oxegen_concentrator_equipment_in_the_past_36_months": toBoolean(row[oxygenConcentratorIndex]),
              "Rechargeable_Spinal_Cord_Simulator_SCS": toBoolean(row[spinalCordStimulatorIndex]),
              "Right_Ventricular_assistive_device_RVAD": toBoolean(row[rvadIndex]),
              "Suction_pump": toBoolean(row[suctionPumpIndex]),
              "Suction_Pump_Machine": toBoolean(row[suctionPumpMachineIndex]),
              "Telephone_Communication_for_Deaf_Hard_of_Hearing": toBoolean(row[deafCommunicationIndex]),
              "Total_artifical_heart_TAH_in_the_past_5_years": toBoolean(row[artificialHeartIndex]),
              "Vaporizer": toBoolean(row[vaporizerIndex]),
              "Rechargeable_Electrical_Neurostimulator_Implant": toBoolean(row[neurostimulatorImplantIndex]),
            };
          const subModuleResult = await insertSubmoduleRecord(submModuleData);
          let submoduleID;

          submoduleID = subModuleResult.data[0].details.id;
          if (!subModuleResult.success) {
            console.error(`Failed to create deal for ${email}:`, subModuleResult.error);
            throw new Error(`Failed to create deal: ${subModuleResult.error}`);
          }


            // Step 4: Insert deal record 
            const dealData = {
                "DOE_ID_Number": number,
                "Deal_Name": firstName + " " + lastName,
                "Latitude": latitude,
                "Longitude": longitude,
                "Customer_Phone": phoneNumber,
                "Layout": { "id": "4909080000146647839" },
                "Customer_State  ": municipality,
                "Customer_Street  ": houseNumber,
                "Customer_Postal_Code  ": zipCode,
                "Customer_City": city,
                "Construction_Year": constructionYear,
                "Is_the_single_dwelling_house_50_yrs_or_older": houseAge,
                "the_house_have_a_flat_or_inclined_pitched_roof": roofType,
                "Is_anyone_eligible_as_an_Energy_Dependent": disability,
                "If_your_medical_equipment_is_not_listed_above": medicalEquipment,
                "Upload_of_Med_Device": "",
                "Assigned": dateAssigned,
                "Contact_Name": contactId,
                "Account_Name": accountId,
                "Tipo_Comercial": "Generac",
                "Stage": "New",
                 "Geographic_Eligibility_Last_Mile_CommunityIndex": geographicEligibility,
                 "Does_the_house_have_roof_type_material_of_Cement_Concrete": housecementConcrete,
                 "Submodule_Comercial": submoduleID
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
          const errorMessage = error instanceof Error ? error.message : 'Error desconocido';
          newImportErrors.push({ rowNumber, email: email || 'N/A', error: errorMessage });
        } finally {
          processed++;
          setImportProgress(Math.round((processed / totalRows) * 100));
          setImportStatus(`Procesados ${processed} de ${totalRows} registros (${successful} exitosos, ${failures} fallidos)`);
        }
      }

      setImportStatus(`Importación completada: ${successful} exitosos, ${failures} fallidos de ${totalRows} registros totales`);
    } catch (error) {
      setImportStatus(`Importación fallida: ${error instanceof Error ? error.message : 'Error desconocido'}`);
      setError(`Importación fallida: ${error instanceof Error ? error.message : 'Error desconocido'}`);
    } finally {
      setImporting(false);
      setImportErrors(newImportErrors);
    }
  };

  const handlePageChange = (_event: React.ChangeEvent<unknown>, value: number) => {
    setPageNumber(value);
    setError(null);
    setValidationErrors([]);
    setHasValidationErrors(false);
    setImportErrors([]);
  };

  const resetData = () => {
    setData([]);
    setPageNumber(1);
    setFileInfo(null);
    setError(null);
    setValidationErrors([]);
    setHasValidationErrors(false);
    setImportErrors([]);
    setImportStatus(null);
  };

  const toggleShowOnlyErrors = () => {
    setPageNumber(1);
    setShowOnlyErrors(prev => !prev);
  };

  const dataRows = useMemo(() => data.length > 1 ? data.slice(1) : [], [data]);

  const rowsToDisplay = useMemo(() => {
    const allItems = dataRows.map((row, index) => ({
        row,
        originalRowNumber: index + 2,
    }));

    if (showOnlyErrors && hasValidationErrors) {
        const errorRowNumbers = new Set(validationErrors.map(e => e.row));
        return allItems.filter(item => errorRowNumbers.has(item.originalRowNumber));
    }
    return allItems;
  }, [dataRows, showOnlyErrors, hasValidationErrors, validationErrors]);

  const pageCount = Math.ceil(rowsToDisplay.length / rowsPerPage);
  const displayRows = rowsToDisplay.slice((pageNumber - 1) * rowsPerPage, pageNumber * rowsPerPage);

  const hasHeaders = data.length > 0;
  const hasData = data.length > 1;

  const headers = data.length > 0 ? data[0] : [];
  const dateAssignedIndex = headers.findIndex(h => String(h).toLowerCase().includes('date assigned'));
  const emailIndex = headers.findIndex(h => String(h).toLowerCase().includes('homeowner') && String(h).toLowerCase().includes('email'));

  // Función para generar y descargar archivo de ejemplo
  const downloadSampleFile = () => {
    const sampleHeaders = [
      'Date Assigned',
      'Number',
      'Name',
      'Latitude:',
      'Longitude:',
      'Phone number:',
      'Alternate phone number:',
      'Homeowner\'s Email:',
      'Municipality',
      'House Number and Street Name',
      'City',
      'Zip Code',
      'House Number and Street Name (Mailing)',
      'City (Mailing)',
      'Municipality (Mailing)',
      'Zip Code (Mailing)',
      'Construction Year (enter 4 digit year ex. 1950)',
      'Is the single dwelling house 50 years of age or older?',
      'Does the house have a flat or inclined/pitched roof?',
      'Does the house have roof type material of Cement/Concrete or Metal/Zinc ?',
      'Geographic Eligibility (Last Mile Community',
      'Is anyone in the household eligible as an Energy Dependent Disability Individual',
      'If your medical equipment is not listed above',
      'Air Conditioner (A/C) for temperature control',
      'Air Purifier',
      'Air mattress for Bed Sores or Alternating Air Pressure Mattress',
      'Asthma therapy machine or Nebulizer',
      'At home dialysis machine',
      'Bilevel positive airway pressure (BiPAP) machine',
      'CPAP, BPAP, APAP or any other Sleep Apnea Machine',
      'Dehumidifier',
      'Dialysis Machine',
      'Electric Crane',
      'Electric Machine for Physical Therapy',
      'Electric Power Lift Recliner',
      'Electric Vital Signs Monitor',
      'Electric bed equipment in the last 13 months',
      'Electric scooter',
      'Electric wheelchair',
      'Energy Dependent Disability Eligibility',
      'Energy Dependent disability observed but homeowner did not consent to pictures',
      'Enteral Feeding Tube Pump Machine / Naso Feeding Machine',
      'Enteral feeding machine',
      'External Defibrillator',
      'FFT Electric Machine',
      'Fan for temperature control',
      'Hearing Aid Pods Rechargeable',
      'Humidifier',
      'Implanted cardiac devices that include left ventricular assistive device(LVAD)',
      'Mechanical Ventilator',
      'Medications that require refrigeration',
      'Oxegen concentrator equipment in the past 36 months',
      'Rechargeable Electrical Neurostimulator Implant',
      'Rechargeable Spinal Cord Simulator (SCS)',
      'Right Ventricular assistive device (RVAD)',
      'Suction Pump Machine',
      'Suction pump',
      'Telephone Communication for Deaf/Hard of Hearing',
      'Total artifical heart (TAH) in the past 5 years',
      'Vaporizer',
      'bi-ventricular assistive device (BIVAD)',
      'intravenous (IV) infusion pump'
    ];

    const sampleData = [
      sampleHeaders,
      [
        '01/15/2024',
        '001',
        'Juan Pérez',
        '18.4655',
        '-66.1057',
        '787-555-0101',
        '787-555-0102',
        'juan.perez@email.com',
        'San Juan',
        '123 Calle Principal',
        'San Juan',
        '00901',
        '123 Calle Principal',
        'San Juan',
        'San Juan',
        '00901',
        '1980',
        'No',
        'Inclined/pitched',
        'No',
        'No',
        'No',
        'N/A',
        'Yes',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No',
        'No'
      ]
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(sampleData);
    
    // Ajustar ancho de columnas
    const columnWidths = sampleHeaders.map(header => ({ wch: Math.max(header.length, 15) }));
    worksheet['!cols'] = columnWidths;
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sample Data');
    XLSX.writeFile(workbook, 'archivo_ejemplo_columnas.xlsx');
  };

  const handleOpenInfoDialog = () => {
    setInfoDialogOpen(true);
  };

  const handleCloseInfoDialog = () => {
    setInfoDialogOpen(false);
  };

  return (
    <Paper elevation={3} sx={{ width: '100%', overflowX: 'auto', p: { xs: 1, sm: 2 } }}>
        <input
          type="file"
          accept=".csv,.xlsx"
          onChange={handleFileUpload}
          style={{ display: "none" }}
          id="file-upload"
          disabled={loading || importing}
        />
      <Box display="flex" flexDirection="column" alignItems="center" mb={2} gap={1}>
        {/* Fila de acciones principales */}
        <Box display="flex" alignItems="center" justifyContent="center" gap={2} flexWrap="wrap">
        <label htmlFor="file-upload">
          <Button
            variant="contained"
            component="span"
            color="primary"
            startIcon={<CloudUploadIcon />}
            disabled={loading || importing}
          >
              Subir Archivo
          </Button>
        </label>

        {hasData && (
          <>
            <Button
              variant="contained"
              color="success"
              startIcon={<UploadIcon />}
              onClick={handleImport}
                disabled={loading || importing || hasValidationErrors}
            >
                Importar
            </Button>
            <Button
                variant="text"
              color="error"
              startIcon={<DeleteIcon />}
              onClick={resetData}
              disabled={loading || importing}
            >
                Limpiar
            </Button>
          </>
        )}
        </Box>
        
        {/* Fila de acciones de ayuda */}
        <Box display="flex" alignItems="center" justifyContent="center" gap={2} flexWrap="wrap" mt={1}>
          <Button
            variant="text"
            color="info"
            startIcon={<DownloadIcon />}
            onClick={downloadSampleFile}
            disabled={loading || importing}
            sx={{ textTransform: 'none' }}
          >
            Descargar Archivo de Ejemplo
          </Button>

          <Button
            variant="text"
            color="info"
            startIcon={<InfoIcon />}
            onClick={handleOpenInfoDialog}
            disabled={loading || importing}
            sx={{ textTransform: 'none' }}
          >
            Información del Proceso
          </Button>
        </Box>
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

      {!importing && importStatus && !hasValidationErrors && (
        <Alert 
          severity={importErrors.length > 0 ? "warning" : "success"} 
          sx={{ mb: 2 }}
        >
          {importStatus}
        </Alert>
      )}

      {!importing && importErrors.length > 0 && (
        <Alert severity="error" sx={{ mb: 2 }}>
          <Typography variant="h6" gutterBottom>
            Se encontraron errores durante la importación
          </Typography>
          <Typography variant="body2" gutterBottom>
            Los siguientes registros no pudieron ser importados:
          </Typography>
          <Box sx={{ maxHeight: '200px', overflowY: 'auto', mt: 1, bgcolor: 'rgba(255, 0, 0, 0.05)', p: 1, borderRadius: 1 }}>
            {importErrors.map((error, index) => (
              <Box key={index} sx={{ mb: 1 }}>
                <Typography variant="body2" sx={{ fontWeight: 'bold' }}>
                  Fila {error.rowNumber}: {error.error}
                </Typography>
                {error.email && (
                  <Typography variant="caption" sx={{ color: 'text.secondary' }}>
                    Email: {error.email}
                  </Typography>
                )}
              </Box>
            ))}
          </Box>
        </Alert>
      )}

      {loading && (
        <Box display="flex" justifyContent="center" my={4}>
          <CircularProgress />
        </Box>
      )}

      {fileInfo && (
        <Box mb={2} mt={2}>
          <Typography variant="subtitle1" gutterBottom>
            Archivo cargado: {fileInfo.name}
          </Typography>
          <Box display="flex" justifyContent="space-between" alignItems="center" flexWrap="wrap" gap={1}>
          <Box display="flex" gap={1} flexWrap="wrap">
              <Chip label={`Tipo: ${fileInfo.type}`} color="primary" size="small" />
              <Chip label={`Tamaño: ${formatFileSize(fileInfo.size)}`} color="primary" size="small" />
              <Chip label={`Filas: ${fileInfo.rows}`} color="primary" size="small" />
              <Chip label={`Columnas: ${data[0]?.length || 0}`} color="info" size="small" />
              {hasValidationErrors && (
                <Tooltip title={showOnlyErrors ? "Mostrar todos" : "Mostrar solo errores"}>
                  <Chip 
                    label={`${validationErrors.length} errores de validación`} 
                    color="warning" 
                    size="small" 
                    icon={showOnlyErrors ? <FilterListOffIcon /> : <FilterListIcon />}
                    onClick={toggleShowOnlyErrors}
                    onDelete={toggleShowOnlyErrors}
                  />
                </Tooltip>
              )}
          </Box>

            {hasValidationErrors && (
              <Button
                variant="outlined"
                color="warning"
                size="small"
                startIcon={<PlaylistRemoveIcon />}
                onClick={handleProcessErrors}
              >
                Descargar y Eliminar Errores
              </Button>
            )}
          </Box>
          <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>
            💡 Las celdas vacías se muestran como <strong>[Vacío]</strong> para mantener la alineación correcta de las columnas.
          </Typography>
        </Box>
      )}

      {!loading && !hasData && !error && (
        <Box textAlign="center" my={6}>
          <Typography color="textSecondary" gutterBottom>
            No hay datos para mostrar. Por favor, suba un archivo CSV o XLSX.
          </Typography>
          <Alert severity="info" sx={{ mt: 2, maxWidth: '600px', mx: 'auto' }}>
            <Typography variant="body2" gutterBottom>
              <strong>¿No estás seguro del formato requerido?</strong>
            </Typography>
            <Typography variant="body2">
              Descarga el archivo de ejemplo para ver exactamente qué columnas necesitas incluir y en qué orden.
              El archivo incluye todas las columnas requeridas con datos de ejemplo.
            </Typography>
          </Alert>
        </Box>
      )}

      {hasData && (
        <>
          <TableContainer component={Paper} sx={{ maxHeight: "70vh" }}>
            <Table size="small" stickyHeader>
              <TableHead>
                <TableRow>
                  <TableCell
                    sx={{ 
                      fontWeight: 'bold', 
                      bgcolor: 'primary.main',
                      color: 'common.black',
                      width: '60px',
                      textAlign: 'center'
                    }}
                  >
                    #
                  </TableCell>
                  {data[0].map((header, index) => (
                    <TableCell
                      key={index}
                      sx={{ 
                        fontWeight: 'bold', 
                        bgcolor: 'primary.main',
                        color: 'common.black',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      {String(header || `Column ${index + 1}`)}
                    </TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {displayRows.map(({ row, originalRowNumber }) => {
                  const actualRowNumber = originalRowNumber;
                  const rowError = validationErrors.find(e => e.row === actualRowNumber);

                  return (
                    <TableRow 
                      key={actualRowNumber}
                      hover 
                      sx={{ 
                        '&:hover': { backgroundColor: 'action.hover' },
                        ...(rowError && { bgcolor: 'rgba(211, 47, 47, 0.1)' })
                      }}
                    >
                      <TableCell
                        sx={{ 
                          fontWeight: 'bold', 
                          bgcolor: 'grey.200',
                          textAlign: 'center',
                          borderRight: '1px solid',
                          borderColor: 'divider'
                        }}
                      >
                        {actualRowNumber}
                      </TableCell>
                      {row.map((cell, cellIndex) => {
                        const isEmailErrorCell = cellIndex === emailIndex && rowError;

                        return (
                          <TableCell key={`${actualRowNumber}-${cellIndex}`}>
                            {(() => {
                              if (cellIndex === dateAssignedIndex && typeof cell === 'number' && cell > 1) {
                                return XLSX.SSF.format('m/d/yy', cell);
                              }
                              
                              const isEmpty = cell === null || cell === undefined || cell === '';

                              if (isEmailErrorCell) {
                                const errorContent = (
                                  <Typography variant="body2" sx={{ color: 'error.main', fontWeight: 'bold', display: 'inline-block' }}>
                                    {isEmpty ? '[Vacío]' : String(cell)}
                                  </Typography>
                                );

                                return (
                                  <Tooltip title={rowError.message} arrow>
                                    {errorContent}
                                  </Tooltip>
                                );
                              }

                              if (isEmpty) {
                                return (
                                  <Typography 
                                    variant="body2" 
                                    sx={{ 
                                      color: 'text.secondary', 
                                      fontStyle: 'italic',
                                      bgcolor: 'grey.100',
                                      px: 1,
                                      py: 0.5,
                                      borderRadius: 1,
                                      display: 'inline-block'
                                    }}
                                  >
                                    [Vacío]
                                  </Typography>
                                );
                              }
                              return String(cell);
                            })()}
                          </TableCell>
                        );
                      })}
                  </TableRow>
                  );
                })}
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

      {/* Diálogo de Información del Proceso */}
      <Dialog
        open={infoDialogOpen}
        onClose={handleCloseInfoDialog}
        maxWidth="md"
        fullWidth
      >
        <DialogTitle sx={{ 
          bgcolor: 'primary.main', 
          color: 'common.black',
          display: 'flex',
          alignItems: 'center',
          gap: 1
        }}>
          <InfoIcon />
          Información del Proceso de Importación
        </DialogTitle>
        
        <DialogContent sx={{ pt: 3 }}>
          <ProcessFlowDiagram />
        </DialogContent>

        <DialogActions>
          <Button onClick={handleCloseInfoDialog} color="primary">
            Entendido
          </Button>
        </DialogActions>
      </Dialog>
    </Paper>
  );
};

const ProcessFlowDiagram = () => {
  const theme = useTheme();
  const steps = [
    { primary: "1. Validación", secondary: "Se verifican los datos del archivo." },
    { primary: "2. Búsqueda", secondary: "Se busca el contacto por email." },
    { primary: "3. Account", secondary: "Se crea la cuenta si no existe." },
    { primary: "4. Contacto", secondary: "Se crea o actualiza el contacto." },
    { primary: "5. Submódulo", secondary: "Se registran los datos de equipos." },
    { primary: "6. Deal", secondary: "Se crea la oportunidad final." }
  ];

  return (
    <>
      <Typography variant="h6" gutterBottom color="primary">
        ¿Qué hace este proceso?
      </Typography>
      
      <Typography variant="body1" paragraph>
        Este sistema está diseñado para generar registros en el <strong>módulo de Deals tipo Generac</strong>&nbsp;en Zoho CRM. El proceso automatiza la creación de múltiples entidades relacionadas.
      </Typography>

      <Divider sx={{ my: 2 }} />

      <Typography variant="h6" gutterBottom color="primary">
        Entidades que se crean automáticamente:
      </Typography>

      <List>
        <ListItem>
          <ListItemIcon>
            <AccountCircleIcon color="primary" />
          </ListItemIcon>
          <ListItemText
            primary="Account (Cuenta)"
            secondary="Se crea una cuenta para cada propietario de vivienda con la información básica de contacto."
          />
        </ListItem>

        <ListItem>
          <ListItemIcon>
            <BusinessIcon color="primary" />
          </ListItemIcon>
          <ListItemText
            primary="Contact (Contacto)"
            secondary="Se crea un contacto individual con toda la información personal y de ubicación del propietario."
          />
        </ListItem>

        <ListItem>
          <ListItemIcon>
            <CheckCircleIcon color="primary" />
          </ListItemIcon>
          <ListItemText
            primary="Deal (Oportunidad)"
            secondary="Se crea una oportunidad de venta tipo 'Generac' con toda la información del proyecto y la vivienda."
          />
        </ListItem>

        <ListItem>
          <ListItemIcon>
            <InfoIcon color="primary" />
          </ListItemIcon>
          <ListItemText
            primary="Submódulo Generac"
            secondary="Se crea un registro en el submódulo específico de Generac con toda la información de equipos médicos y elegibilidad."
          />
        </ListItem>
      </List>

      <Divider sx={{ my: 2 }} />

      <Typography variant="h6" gutterBottom color="error">
        ⚠️ Importante: Validación de Emails
      </Typography>

      <Box sx={{ 
        bgcolor: 'background.paper',
        p: 2, 
        borderRadius: 1,
        border: '2px solid',
        borderColor: 'error.main'
      }}>
        <Typography variant="body1" paragraph>
          <strong>El email es el identificador primario</strong> para la creación de contactos. 
          Es fundamental que:
        </Typography>
        
        <List dense>
          <ListItem>
            <ListItemIcon>
              <EmailIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary="Todos los registros tengan un email válido"
              secondary="No se pueden crear contactos sin email"
            />
          </ListItem>
          
          <ListItem>
            <ListItemIcon>
              <WarningIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary="Los emails sean únicos"
              secondary="No se pueden duplicar contactos con el mismo email"
            />
          </ListItem>
          
          <ListItem>
            <ListItemIcon>
              <CheckCircleIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary="El formato sea correcto"
              secondary="Ejemplo: usuario@dominio.com"
            />
          </ListItem>
        </List>
      </Box>

      <Divider sx={{ my: 2 }} />

      <Typography variant="h6" gutterBottom color="primary">
        Flujo del Proceso:
      </Typography>

      <Box sx={{ display: 'flex', flexDirection: 'row', alignItems: 'stretch', justifyContent: 'center', gap: 1, mt: 2, flexWrap: 'wrap' }}>
        {steps.map((step, index) => (
          <React.Fragment key={index}>
            <Paper 
              elevation={2} 
              sx={{ 
                p: 2, 
                flex: 1,
                minWidth: '150px',
                maxWidth: '200px',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'center',
                textAlign: 'center', 
                borderTop: `4px solid ${theme.palette.primary.main}` 
              }}
            >
              <Typography variant="subtitle1" sx={{ fontWeight: 'bold' }}>{step.primary}</Typography>
              <Typography variant="body2" color="text.secondary">{step.secondary}</Typography>
            </Paper>
            {index < steps.length - 1 && (
              <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'center', alignSelf: 'center' }}>
                <ArrowForwardIcon color="primary" sx={{ fontSize: '2rem' }} />
              </Box>
            )}
          </React.Fragment>
        ))}
      </Box>

      <Alert severity="info" sx={{ mt: 3 }}>
        <Typography variant="body2">
          <strong>Nota:</strong> Si un contacto ya existe, el sistema lo reutiliza para asociarle un nuevo Deal, en lugar de crear un duplicado. Esto garantiza la integridad de los datos.
        </Typography>
      </Alert>
    </>
  );
};

export default UploadExcel;
