import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { Box, Container, Typography, Paper, Button, TextField, Grid, Alert,
  LinearProgress, Chip, Card, CardContent, CardActions, Stepper, Step,
  StepLabel,  List,   ListItem, ListItemIcon, ListItemText, IconButton,
  Fade, Dialog, DialogTitle, DialogContent, DialogActions, Divider
} from '@mui/material';
import {
  CloudUpload, Description, TableChart, Settings, PlayArrow, CheckCircle,
  Error as ErrorIcon, Info, Delete, Clear, GetApp, FileCopy, FolderOpen
} from '@mui/icons-material';
import { styled } from '@mui/material/styles';

// Styled components matching docToPdf.js
const GradientHeader = styled(Paper)(({ theme }) => ({
  background: 'linear-gradient(135deg, #d4262fff 0%, #af3194ff 100%)',
  color: 'white',
  padding: theme.spacing(4),
  borderRadius: 0,
  textAlign: 'center'
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
  const [activeStep, setActiveStep] = useState(0);
  const [showDialog, setShowDialog] = useState(false);
  const [dialogContent, setDialogContent] = useState({});

  const steps = ['Upload Files', 'Configure Settings', 'Merge Documents'];

  const handleFileSelect = (event, fileType) => {
    const file = event.target.files[0];
    setFiles(prev => ({
      ...prev,
      [fileType]: file
    }));
    
    // Move to next step if all files are selected
    if (canProceedToSettings()) {
      setActiveStep(1);
    }
  };

  const removeFile = (fileType) => {
    setFiles(prev => ({
      ...prev,
      [fileType]: null
    }));
    
    // Go back to step 0 if files are removed
    if (!canProceedToSettings()) {
      setActiveStep(0);
    }
  };

  const clearAllFiles = () => {
    setFiles({
      refTemplate: null,
      refData: null,
      docTemplate: null,
      docData: null
    });
    setStatus(null);
    setActiveStep(0);
  };

  const getFileStatus = (fileType) => {
    const file = files[fileType];
    if (file) {
      return { text: `✓ ${file.name}`, uploaded: true };
    }
    return { text: 'No file selected', uploaded: false };
  };

  const canProceedToSettings = () => {
    return Object.values(files).every(file => file !== null);
  };

  const canMerge = () => {
    return canProceedToSettings() && !isProcessing;
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
    setActiveStep(2);
    
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
        
        setDialogContent({
          title: 'Merge Complete',
          message: 'Documents merged successfully!',
          detail: `Successfully merged ${allFiles.length} documents into a ZIP file.\n\nThe download should start automatically.`,
          onConfirm: () => {
            setShowDialog(false);
            reset();
          }
        });
        setShowDialog(true);
        
        setIsProcessing(false);
      }, 500);
      
    } catch (error) {
      console.error('Merge error:', error);
      showStatus(`❌ Error: ${error.message}`, 'error');
      setIsProcessing(false);
    }
  };

  const reset = () => {
    setFiles({
      refTemplate: null,
      refData: null,
      docTemplate: null,
      docData: null
    });
    setStatus(null);
    setActiveStep(0);
    setIsProcessing(false);
  };

  const fileTypes = [
    { key: 'refTemplate', label: 'Reference Template', accept: '.docx', icon: Description, description: 'Word template for references' },
    { key: 'refData', label: 'Reference Data', accept: '.xlsx,.xls', icon: TableChart, description: 'Excel data for references' },
    { key: 'docTemplate', label: 'Document Template', accept: '.docx', icon: Description, description: 'Word template for documents' },
    { key: 'docData', label: 'Document Data', accept: '.xlsx,.xls', icon: TableChart, description: 'Excel data for documents' }
  ];

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Paper elevation={3} sx={{ overflow: 'hidden' }}>
        {/* Header */}
        <GradientHeader>
          <Typography variant="h4" component="h1" gutterBottom>
            Document Merge Tool
          </Typography>
          <Typography variant="subtitle1">
            Merge Excel data with Word document templates seamlessly. Use merge fields like {`{{fieldname}}`} in your Word documents.
          </Typography>
        </GradientHeader>

        {/* Stepper */}
        <Box sx={{ p: 3, backgroundColor: 'grey.50' }}>
          <Stepper activeStep={activeStep} alternativeLabel>
            {steps.map((label, index) => (
              <Step key={label}>
                <StepLabel
                  StepIconComponent={({ active, completed }) => {
                    const icons = [CloudUpload, Settings, PlayArrow];
                    const Icon = icons[index];
                    return (
                      <Icon 
                        sx={{ 
                          color: active ? 'primary.main' : completed ? 'success.main' : 'grey.400',
                          fontSize: 24
                        }} 
                      />
                    );
                  }}
                >
                  {label}
                </StepLabel>
              </Step>
            ))}
          </Stepper>
        </Box>

        <Box sx={{ p: 3 }}>
          {/* File Upload Step */}
          {activeStep === 0 && (
            <Fade in timeout={500}>
              <Box>
                <Typography variant="h6" gutterBottom sx={{ mb: 3 }}>
                  Upload Required Files
                </Typography>
                
                <Grid container spacing={3}>
                  {fileTypes.map((fileType) => {
                    const fileStatus = getFileStatus(fileType.key);
                    const Icon = fileType.icon;
                    
                    return (
                      <Grid item xs={12} md={6} key={fileType.key}>
                        <Card sx={{ height: '100%' }}>
                          <CardContent>
                            <Box sx={{ display: 'flex', alignItems: 'center', mb: 2 }}>
                              <Icon sx={{ mr: 1, color: 'primary.main' }} />
                              <Typography variant="h6">
                                {fileType.label}
                              </Typography>
                            </Box>
                            
                            <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>
                              {fileType.description}
                            </Typography>

                            <input
                              accept={fileType.accept}
                              style={{ display: 'none' }}
                              id={`${fileType.key}-upload`}
                              type="file"
                              onChange={(e) => handleFileSelect(e, fileType.key)}
                              disabled={isProcessing}
                            />
                            <label htmlFor={`${fileType.key}-upload`}>
                              <Button
                                component="span"
                                variant="outlined"
                                startIcon={<CloudUpload />}
                                fullWidth
                                disabled={isProcessing}
                                sx={{ mb: 2 }}
                              >
                                Choose File
                              </Button>
                            </label>

                            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                              <Chip
                                label={fileStatus.text}
                                color={fileStatus.uploaded ? 'success' : 'default'}
                                variant={fileStatus.uploaded ? 'filled' : 'outlined'}
                                size="small"
                              />
                              
                              {files[fileType.key] && (
                                <IconButton
                                  size="small"
                                  onClick={() => removeFile(fileType.key)}
                                  disabled={isProcessing}
                                  color="error"
                                >
                                  <Delete />
                                </IconButton>
                              )}
                            </Box>
                          </CardContent>
                        </Card>
                      </Grid>
                    );
                  })}
                </Grid>

                <Alert severity="info" sx={{ mt: 3 }}>
                  <Typography variant="body2">
                    💡 Upload all 4 files to proceed • Templates should contain merge fields like {`{{fieldname}}`} • Excel files provide the data for merging
                  </Typography>
                </Alert>

                {canProceedToSettings() && (
                  <Box sx={{ textAlign: 'center', mt: 3 }}>
                    <Button
                      variant="contained"
                      onClick={() => setActiveStep(1)}
                      disabled={isProcessing}
                      startIcon={<Settings />}
                      size="large"
                    >
                      Configure Settings
                    </Button>
                  </Box>
                )}
              </Box>
            </Fade>
          )}

          {/* Settings Step */}
          {activeStep === 1 && (
            <Fade in timeout={500}>
              <Grid container spacing={3}>
                <Grid item xs={12} lg={8}>
                  <Card>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        Uploaded Files
                      </Typography>
                      
                      <List>
                        {fileTypes.map((fileType) => {
                          const file = files[fileType.key];
                          const Icon = fileType.icon;
                          
                          return (
                            <ListItem
                              key={fileType.key}
                              sx={{
                                border: 1,
                                borderColor: 'grey.200',
                                borderRadius: 1,
                                mb: 1,
                                backgroundColor: 'grey.50'
                              }}
                            >
                              <ListItemIcon>
                                <Icon color="primary" />
                              </ListItemIcon>
                              <ListItemText
                                primary={fileType.label}
                                secondary={file ? file.name : 'No file selected'}
                              />
                              {file && (
                                <IconButton
                                  onClick={() => removeFile(fileType.key)}
                                  disabled={isProcessing}
                                  color="error"
                                >
                                  <Delete />
                                </IconButton>
                              )}
                            </ListItem>
                          );
                        })}
                      </List>
                    </CardContent>
                  </Card>
                </Grid>

                <Grid item xs={12} lg={4}>
                  <Card sx={{ height: 'fit-content' }}>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        Output Settings
                      </Typography>
                      
                      <TextField
                        fullWidth
                        label="Output Folder Name"
                        value={outputLocation}
                        onChange={(e) => setOutputLocation(e.target.value)}
                        placeholder="Enter output folder name"
                        variant="outlined"
                        disabled={isProcessing}
                        sx={{ mb: 3 }}
                      />

                      <Alert severity="success" sx={{ mb: 3 }}>
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                          <GetApp />
                          <Box>
                            <Typography variant="subtitle2">
                              ZIP Download
                            </Typography>
                            <Typography variant="body2">
                              All merged documents will be packaged in a ZIP file
                            </Typography>
                          </Box>
                        </Box>
                      </Alert>

                      <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                        <Button
                          fullWidth
                          variant="contained"
                          color="primary"
                          onClick={startMerge}
                          disabled={!canMerge()}
                          startIcon={<PlayArrow />}
                          size="large"
                        >
                          {isProcessing ? 'Processing...' : 'Start Merge Process'}
                        </Button>
                        
                        <Button
                          fullWidth
                          variant="outlined"
                          onClick={() => setActiveStep(0)}
                          disabled={isProcessing}
                        >
                          Back to Upload
                        </Button>
                        
                        <Button
                          fullWidth
                          variant="outlined"
                          color="error"
                          onClick={clearAllFiles}
                          disabled={isProcessing}
                          startIcon={<Clear />}
                        >
                          Clear All Files
                        </Button>
                      </Box>
                    </CardContent>
                  </Card>
                </Grid>
              </Grid>
            </Fade>
          )}

          {/* Merge Progress Step */}
          {activeStep === 2 && (
            <Fade in timeout={500}>
              <Box>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 3 }}>
                  <Typography variant="h6">
                    Merge Progress
                  </Typography>
                  <Button 
                    onClick={() => setActiveStep(1)}
                    variant="outlined"
                    disabled={isProcessing}
                  >
                    Back to Settings
                  </Button>
                </Box>

                <Card sx={{ mb: 3 }}>
                  <CardContent>
                    <Typography variant="subtitle1" gutterBottom>
                      Merge Summary
                    </Typography>
                    <Grid container spacing={2}>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Files to Process:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {Object.values(files).filter(f => f !== null).length}
                        </Typography>
                      </Grid>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Output Folder:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {outputLocation}
                        </Typography>
                      </Grid>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Status:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {isProcessing ? 'Processing...' : 'Ready'}
                        </Typography>
                      </Grid>
                    </Grid>
                  </CardContent>
                </Card>

                {/* Status Section */}
                {status && (
                  <Card>
                    <CardContent>
                      <Box sx={{ display: 'flex', alignItems: 'center', gap: 2, mb: 2 }}>
                        {status.type === 'success' && <CheckCircle color="success" />}
                        {status.type === 'error' && <ErrorIcon color="error" />}
                        {status.type === 'processing' && <Info color="info" />}
                        
                        <Typography variant="h6" sx={{ 
                          color: status.type === 'success' ? 'success.main' : 
                                 status.type === 'error' ? 'error.main' : 'info.main'
                        }}>
                          {status.message}
                        </Typography>
                      </Box>
                      
                      {status.progress !== null && (
                        <Box sx={{ width: '100%' }}>
                          <LinearProgress 
                            variant="determinate" 
                            value={status.progress} 
                            sx={{ height: 8, borderRadius: 4 }}
                          />
                          <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>
                            {status.progress}% Complete
                          </Typography>
                        </Box>
                      )}
                    </CardContent>
                  </Card>
                )}

                <Box sx={{ textAlign: 'center', mt: 3 }}>
                  <Button 
                    onClick={reset}
                    variant="outlined"
                    startIcon={<CloudUpload />}
                    disabled={isProcessing}
                  >
                    Merge More Files
                  </Button>
                </Box>
              </Box>
            </Fade>
          )}
        </Box>
      </Paper>

      {/* Dialog for completion confirmation */}
      <Dialog open={showDialog} onClose={() => setShowDialog(false)}>
        <DialogTitle>{dialogContent.title}</DialogTitle>
        <DialogContent>
          <Typography variant="body1" gutterBottom>
            {dialogContent.message}
          </Typography>
          <Typography variant="body2" color="text.secondary">
            {dialogContent.detail}
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setShowDialog(false)}>
            Cancel
          </Button>
          <Button onClick={dialogContent.onConfirm} variant="contained" autoFocus>
            OK
          </Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
};

export default FileMerge;