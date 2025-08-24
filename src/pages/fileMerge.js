import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {
  Box,
  Container,
  Typography,
  Paper,
  Button,
  TextField,
  Grid,
  Alert,
  LinearProgress,
  Chip,
  Card,
  CardContent,
  styled
} from '@mui/material';
import {
  CloudUpload,
  Description,
  TableChart,
  Settings,
  PlayArrow
} from '@mui/icons-material';

// Styled components
const GradientBox = styled(Box)({
  background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
  minHeight: '100vh',
  padding: '20px',
});

const GlassContainer = styled(Container)({
  maxWidth: '1200px',
  background: 'rgba(255, 255, 255, 0.95)',
  backdropFilter: 'blur(10px)',
  borderRadius: '20px',
  boxShadow: '0 20px 40px rgba(0, 0, 0, 0.1)',
  overflow: 'hidden',
  padding: 0,
});

const HeaderBox = styled(Box)({
  background: 'linear-gradient(45deg, #667eea, #764ba2)',
  color: 'white',
  padding: '30px',
  textAlign: 'center',
});

const SectionCard = styled(Card)({
  background: 'white',
  borderRadius: '15px',
  marginBottom: '30px',
  boxShadow: '0 5px 15px rgba(0, 0, 0, 0.08)',
  border: '1px solid #f0f0f0',
});

const UploadButton = styled(Button)({
  background: 'linear-gradient(45deg, #667eea, #764ba2)',
  color: 'white',
  padding: '15px 20px',
  borderRadius: '10px',
  fontWeight: 500,
  minHeight: '60px',
  width: '100%',
  '&:hover': {
    background: 'linear-gradient(45deg, #5a6fd8, #6a4190)',
    transform: 'translateY(-2px)',
    boxShadow: '0 5px 15px rgba(102, 126, 234, 0.4)',
  },
  transition: 'all 0.3s ease',
});

const MergeButton = styled(Button)(({ disabled }) => ({
  background: disabled 
    ? '#6c757d' 
    : 'linear-gradient(45deg, #28a745, #20c997)',
  color: 'white',
  padding: '15px 30px',
  borderRadius: '10px',
  fontSize: '1.1rem',
  fontWeight: 600,
  width: '100%',
  '&:hover': disabled ? {} : {
    transform: 'translateY(-2px)',
    boxShadow: '0 5px 15px rgba(40, 167, 69, 0.4)',
  },
  '&:disabled': {
    background: '#6c757d',
    cursor: 'not-allowed',
  },
  transition: 'all 0.3s ease',
}));

const FileMerge = () => {
  const [files, setFiles] = useState({
    refTemplate: null,
    refData: null,
    docTemplate: null,
    docData: null
  });
  
  const [outputLocation, setOutputLocation] = useState('merged_documents');
  const [status, setStatus] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileSelect = (event, fileType) => {
    const file = event.target.files[0];
    setFiles(prev => ({
      ...prev,
      [fileType]: file
    }));
  };

  const getFileStatus = (fileType) => {
    const file = files[fileType];
    if (file) {
      return { text: `✓ ${file.name}`, className: 'file-uploaded' };
    }
    return { text: 'No file selected', className: 'file-pending' };
  };

  const canMerge = () => {
    return Object.values(files).every(file => file !== null) && !isProcessing;
  };

  const showStatus = (message, type = 'processing', progress = null) => {
    setStatus({ message, type, progress });
  };

  const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {type: 'array'});
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet);
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read Excel file'));
      reader.readAsArrayBuffer(file);
    });
  };

  const readDocxTemplate = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          resolve(e.target.result);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read DOCX file'));
      reader.readAsArrayBuffer(file);
    });
  };

  const mergeDocument = async (templateBuffer, data, startIndex, outputFolder) => {
    const mergedDocs = [];
    
    for (let i = 0; i < data.length; i++) {
      try {
        // Load the template
        const zip = new JSZip();
        await zip.loadAsync(templateBuffer);
        
        // Read the main document content
        const docXml = await zip.file("word/document.xml").async("string");
        
        // Convert data to string values like the original Python script
        const mergeData = {};
        Object.keys(data[i]).forEach(key => {
          mergeData[key] = String(data[i][key] || '');
        });
        
        // Replace merge fields in the document
        let mergedXml = docXml;
        Object.keys(mergeData).forEach(key => {
          // Handle various merge field formats
          const patterns = [
            new RegExp(`{{${key}}}`, 'g'),
            new RegExp(`«${key}»`, 'g'),
            new RegExp(`<<${key}>>`, 'g'),
            new RegExp(`\\$\\{${key}\\}`, 'g')
          ];
          
          patterns.forEach(pattern => {
            mergedXml = mergedXml.replace(pattern, mergeData[key]);
          });
          
          // Also handle Word merge fields format
          const wordMergePattern = new RegExp(`<w:fldSimple[^>]*w:instr="[^"]*MERGEFIELD\\s+${key}\\s*[^"]*"[^>]*>.*?</w:fldSimple>`, 'g');
          mergedXml = mergedXml.replace(wordMergePattern, `<w:r><w:t>${mergeData[key]}</w:t></w:r>`);
        });
        
        // Update the document with merged content
        zip.file("word/document.xml", mergedXml);
        
        // Generate the merged document
        const buf = await zip.generateAsync({type: 'arraybuffer'});
        
        const fileName = `${outputFolder}/${startIndex + i}.docx`;
        mergedDocs.push({
          name: fileName,
          data: buf
        });
      } catch (error) {
        console.error(`Error merging document ${i}:`, error);
        throw new Error(`Failed to merge document ${i + 1}: ${error.message}`);
      }
    }
    
    return mergedDocs;
  };

  const downloadFiles = async (files, outputFolder) => {
    // Create a new zip file containing all merged documents
    const zip = new JSZip();
    
    // Add all files to the zip
    files.forEach(file => {
      const filename = file.name.split('/').pop(); // Get just the filename
      zip.file(filename, file.data);
    });
    
    // Generate the zip file
    const zipBlob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: {
        level: 6
      }
    });
    
    // Download the zip file
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${outputFolder}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const startMerge = async () => {
    setIsProcessing(true);
    
    try {
      showStatus('Starting merge process...', 'processing', 0);
      
      // Read all files
      showStatus('Reading reference files...', 'processing', 20);
      const [refTemplateBuffer, refData] = await Promise.all([
        readDocxTemplate(files.refTemplate),
        readExcelFile(files.refData)
      ]);
      
      showStatus('Reading document files...', 'processing', 40);
      const [docTemplateBuffer, docData] = await Promise.all([
        readDocxTemplate(files.docTemplate),
        readExcelFile(files.docData)
      ]);
      
      // Merge reference documents (starting from index 1)
      showStatus('Merging reference documents...', 'processing', 60);
      const refMerged = await mergeDocument(refTemplateBuffer, refData, 1, outputLocation);
      
      // Merge document files (starting from index 2, but adjusted for ref count)
      showStatus('Merging document files...', 'processing', 80);
      const docMerged = await mergeDocument(docTemplateBuffer, docData, refData.length + 1, outputLocation);
      
      // Download all files
      showStatus('Preparing downloads...', 'processing', 95);
      const allFiles = [...refMerged, ...docMerged];
      
      setTimeout(async () => {
        await downloadFiles(allFiles, outputLocation);
        showStatus(`✅ Successfully merged ${allFiles.length} documents! Zip file download started.`, 'success', 100);
        setIsProcessing(false);
      }, 500);
      
    } catch (error) {
      console.error('Merge error:', error);
      showStatus(`❌ Error: ${error.message}`, 'error');
      setIsProcessing(false);
    }
  };

  return (
    <GradientBox>
      <GlassContainer>
        <HeaderBox>
          <Typography variant="h3" component="h1" sx={{ fontWeight: 700, mb: 1 }}>
            Document Merge Tool
          </Typography>
          <Typography variant="h6" sx={{ opacity: 0.9 }}>
            Merge Excel data with Word document templates seamlessly. Use merge fields like {`{{fieldname}}`} in your Word documents.
          </Typography>
        </HeaderBox>
        
        <Box sx={{ p: 5 }}>
          {/* Reference Files Section */}
          <SectionCard>
            <CardContent sx={{ p: 4 }}>
              <Typography variant="h5" component="h2" sx={{ mb: 3, display: 'flex', alignItems: 'center', gap: 1 }}>
                <Description /> Reference Files
              </Typography>
              <Grid container spacing={3}>
                <Grid item xs={12} md={6}>
                  <input
                    accept=".docx"
                    style={{ display: 'none' }}
                    id="ref-template-upload"
                    type="file"
                    onChange={(e) => handleFileSelect(e, 'refTemplate')}
                  />
                  <label htmlFor="ref-template-upload">
                    <UploadButton component="span" startIcon={<CloudUpload />}>
                      Choose Reference Template (.docx)
                    </UploadButton>
                  </label>
                  <Box sx={{ mt: 1 }}>
                    <Chip
                      label={getFileStatus('refTemplate').text}
                      color={getFileStatus('refTemplate').className === 'file-uploaded' ? 'success' : 'default'}
                      variant={getFileStatus('refTemplate').className === 'file-uploaded' ? 'filled' : 'outlined'}
                      sx={{ width: '100%', py: 1 }}
                    />
                  </Box>
                </Grid>
                <Grid item xs={12} md={6}>
                  <input
                    accept=".xlsx,.xls"
                    style={{ display: 'none' }}
                    id="ref-data-upload"
                    type="file"
                    onChange={(e) => handleFileSelect(e, 'refData')}
                  />
                  <label htmlFor="ref-data-upload">
                    <UploadButton component="span" startIcon={<TableChart />}>
                      Choose Reference Data (.xlsx)
                    </UploadButton>
                  </label>
                  <Box sx={{ mt: 1 }}>
                    <Chip
                      label={getFileStatus('refData').text}
                      color={getFileStatus('refData').className === 'file-uploaded' ? 'success' : 'default'}
                      variant={getFileStatus('refData').className === 'file-uploaded' ? 'filled' : 'outlined'}
                      sx={{ width: '100%', py: 1 }}
                    />
                  </Box>
                </Grid>
              </Grid>
            </CardContent>
          </SectionCard>

          {/* Document Files Section */}
          <SectionCard>
            <CardContent sx={{ p: 4 }}>
              <Typography variant="h5" component="h2" sx={{ mb: 3, display: 'flex', alignItems: 'center', gap: 1 }}>
                <TableChart /> Document Files
              </Typography>
              <Grid container spacing={3}>
                <Grid item xs={12} md={6}>
                  <input
                    accept=".docx"
                    style={{ display: 'none' }}
                    id="doc-template-upload"
                    type="file"
                    onChange={(e) => handleFileSelect(e, 'docTemplate')}
                  />
                  <label htmlFor="doc-template-upload">
                    <UploadButton component="span" startIcon={<CloudUpload />}>
                      Choose Document Template (.docx)
                    </UploadButton>
                  </label>
                  <Box sx={{ mt: 1 }}>
                    <Chip
                      label={getFileStatus('docTemplate').text}
                      color={getFileStatus('docTemplate').className === 'file-uploaded' ? 'success' : 'default'}
                      variant={getFileStatus('docTemplate').className === 'file-uploaded' ? 'filled' : 'outlined'}
                      sx={{ width: '100%', py: 1 }}
                    />
                  </Box>
                </Grid>
                <Grid item xs={12} md={6}>
                  <input
                    accept=".xlsx,.xls"
                    style={{ display: 'none' }}
                    id="doc-data-upload"
                    type="file"
                    onChange={(e) => handleFileSelect(e, 'docData')}
                  />
                  <label htmlFor="doc-data-upload">
                    <UploadButton component="span" startIcon={<TableChart />}>
                      Choose Document Data (.xlsx)
                    </UploadButton>
                  </label>
                  <Box sx={{ mt: 1 }}>
                    <Chip
                      label={getFileStatus('docData').text}
                      color={getFileStatus('docData').className === 'file-uploaded' ? 'success' : 'default'}
                      variant={getFileStatus('docData').className === 'file-uploaded' ? 'filled' : 'outlined'}
                      sx={{ width: '100%', py: 1 }}
                    />
                  </Box>
                </Grid>
              </Grid>
            </CardContent>
          </SectionCard>

          {/* Output Settings Section */}
          <SectionCard>
            <CardContent sx={{ p: 4 }}>
              <Typography variant="h5" component="h2" sx={{ mb: 3, display: 'flex', alignItems: 'center', gap: 1 }}>
                <Settings /> Output Settings
              </Typography>
              <Grid container spacing={3} alignItems="flex-end">
                <Grid item xs={12} md={8}>
                  <TextField
                    fullWidth
                    label="Output Folder Name"
                    value={outputLocation}
                    onChange={(e) => setOutputLocation(e.target.value)}
                    placeholder="Enter output folder name"
                    variant="outlined"
                    sx={{
                      '& .MuiOutlinedInput-root': {
                        borderRadius: '10px',
                        '&:hover fieldset': {
                          borderColor: '#667eea',
                        },
                        '&.Mui-focused fieldset': {
                          borderColor: '#667eea',
                        },
                      },
                      '& .MuiInputLabel-root.Mui-focused': {
                        color: '#667eea',
                      },
                    }}
                  />
                </Grid>
                <Grid item xs={12} md={4}>
                  <MergeButton
                    onClick={startMerge}
                    disabled={!canMerge()}
                    startIcon={<PlayArrow />}
                    size="large"
                  >
                    {isProcessing ? 'Processing...' : 'Start Merge Process'}
                  </MergeButton>
                </Grid>
              </Grid>
            </CardContent>
          </SectionCard>

          {/* Status Section */}
          {status && (
            <Alert
              severity={
                status.type === 'success' ? 'success' :
                status.type === 'error' ? 'error' : 'info'
              }
              sx={{
                borderRadius: '10px',
                fontSize: '1rem',
                '& .MuiAlert-message': {
                  width: '100%',
                },
              }}
            >
              <Typography variant="body1" sx={{ fontWeight: 500 }}>
                {status.message}
              </Typography>
              {status.progress !== null && (
                <LinearProgress
                  variant="determinate"
                  value={status.progress}
                  sx={{
                    mt: 2,
                    height: 8,
                    borderRadius: 4,
                    backgroundColor: 'rgba(0, 0, 0, 0.1)',
                    '& .MuiLinearProgress-bar': {
                      borderRadius: 4,
                      background: 'linear-gradient(45deg, #667eea, #764ba2)',
                    },
                  }}
                />
              )}
            </Alert>
          )}
        </Box>
      </GlassContainer>
    </GradientBox>
  );
};

export default FileMerge;