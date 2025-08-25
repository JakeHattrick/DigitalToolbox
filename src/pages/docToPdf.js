import React, { useState, useRef, useEffect } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  Stepper,
  Step,
  StepLabel,
  Grid,
  Card,
  CardContent,
  CardActions,
  Divider,
  Chip,
  Alert,
  LinearProgress,
  IconButton,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Container,
  Fade,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions
} from '@mui/material';
import {
  CloudUpload,
  Description,
  PictureAsPdf,
  Delete,
  Clear,
  GetApp,
  Merge,
  FileCopy,
  FolderOpen,
  CheckCircle,
  Error as ErrorIcon,
  Info
} from '@mui/icons-material';
import { styled } from '@mui/material/styles';

// Access Electron APIs through the secure preload script
const electronAPI = window.electronAPI;

// Styled components
const GradientHeader = styled(Paper)(({ theme }) => ({
  background: 'linear-gradient(135deg, #af3194ff 0%, #5337d1ff 100%)',
  color: 'white',
  padding: theme.spacing(4),
  borderRadius: 0,
  textAlign: 'center'
}));

const UploadArea = styled(Paper)(({ theme }) => ({
  border: `2px dashed ${theme.palette.grey[300]}`,
  borderRadius: theme.shape.borderRadius,
  padding: theme.spacing(8),
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.3s ease',
  '&:hover': {
    borderColor: theme.palette.primary.main,
    backgroundColor: theme.palette.action.hover
  }
}));

const DocToPdf = () => {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [status, setStatus] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activeStep, setActiveStep] = useState(0);
  const [showDialog, setShowDialog] = useState(false);
  const [dialogContent, setDialogContent] = useState({});

  const steps = ['Select DOCX Files', 'Review Selection', 'Convert & Download'];

  const removeFile = (index) => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
    if (selectedFiles.length === 1) {
      setActiveStep(0);
    }
  };

  const clearAllFiles = () => {
    setSelectedFiles([]);
    setStatus(null);
    setActiveStep(0);
  };

  const showStatus = (message, type = 'processing', progress = null) => {
    console.log(progress);
    setStatus({ message, type, progress });
  };

  const selectFiles = async () => {
    try {
      const result = await electronAPI.selectFiles();
      if (result && !result.canceled && result.filePaths) {
        const fileObjects = result.filePaths
          .filter(path => path.toLowerCase().endsWith('.docx'))
          .map(filePath => ({
            name: electronAPI.path.basename(filePath),
            path: filePath,
            size: 0 // Could get actual file size if needed
          }));
        
        // Remove duplicates and add new files
        const newFiles = fileObjects.filter(file => 
          !selectedFiles.find(f => f.path === file.path)
        );
        setSelectedFiles(prev => [...prev, ...newFiles]);
        
        if (fileObjects.length > 0) {
          setActiveStep(1);
        }
      }
    } catch (error) {
      showStatus('Error selecting files: ' + error.message, 'error');
    }
  };

  useEffect(() => {
    // Subscribe to progress events from the worker
    const unsubscribe = electronAPI?.onDocxToPdfProgress?.((p) => {
      console.log('progress', p);
      showStatus(`Converting documents… ${p}%`, 'processing', p);
    });

    return () => {
      // Clean up the listener when the component unmounts
      if (typeof unsubscribe === 'function') unsubscribe();
    };
  }, []);

  const convertWithWinAX = async () => {
    if (selectedFiles.length === 0) {
      showStatus('No files selected', 'error');
      return;
    }

    // Check if Electron APIs are available
    if (!electronAPI || !electronAPI.convertDocxToPdf) {
      showStatus('Electron APIs not available - this app needs to run in Electron', 'error');
      return;
    }

    setIsProcessing(true);
    setActiveStep(2);
    
    try {
      showStatus('Initializing Word application...', 'processing', 5);

      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      const outputFileName = `merged-documents-${timestamp}.pdf`;
      const downloadsPath = await electronAPI.getDownloadsPath();
      const finalOutputPath = electronAPI.path.join(downloadsPath, outputFileName);

      showStatus('Converting documents with Word...', 'processing', 20);

      // Send conversion request to main process
      const result = await electronAPI.convertDocxToPdf({
        inputFiles: selectedFiles.map(f => f.path),
        outputPath: finalOutputPath,
        mergeDocuments: true
      });

      if (result.success) {
        showStatus(`Successfully converted ${selectedFiles.length} document(s) to PDF!`, 'success', 100);
        
        setDialogContent({
          title: 'Conversion Complete',
          message: 'Documents converted successfully!',
          detail: `PDF saved as: ${outputFileName}\n\nWould you like to open the Downloads folder?`,
          onConfirm: async () => {
            await electronAPI.openDownloadsFolder();
            setShowDialog(false);
          }
        });
        setShowDialog(true);
      } else {
        throw new Error(result.error || 'Unknown conversion error');
      }

    } catch (error) {
      console.error('Conversion error:', error);
      showStatus(`Error during conversion: ${error.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const convertSeparately = async () => {
    if (selectedFiles.length === 0) {
      showStatus('No files selected', 'error');
      return;
    }

    // Check if Electron APIs are available
    if (!electronAPI || !electronAPI.convertDocxToPdf) {
      showStatus('Electron APIs not available - this app needs to run in Electron', 'error');
      return;
    }

    setIsProcessing(true);
    setActiveStep(2);
    
    try {
      showStatus('Initializing Word application...', 'processing', 10);

      const downloadsPath = await electronAPI.getDownloadsPath();

      showStatus('Converting documents individually...', 'processing', 20);

      // Send conversion request to main process for individual files
      const result = await electronAPI.convertDocxToPdf({
        inputFiles: selectedFiles.map(f => f.path),
        outputPath: downloadsPath,
        mergeDocuments: false
      });

      if (result.success) {
        showStatus(`Successfully converted ${selectedFiles.length} document(s) to individual PDFs!`, 'success', 100);
        
        setDialogContent({
          title: 'Conversion Complete',
          message: 'Documents converted successfully!',
          detail: `PDFs saved to Downloads folder\n\nWould you like to open the Downloads folder?`,
          onConfirm: async () => {
            await electronAPI.openDownloadsFolder();
            setShowDialog(false);
          }
        });
        setShowDialog(true);
      } else {
        throw new Error(result.error || 'Unknown conversion error');
      }

    } catch (error) {
      console.error('Conversion error:', error);
      showStatus(`Error during conversion: ${error.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const reset = () => {
    setSelectedFiles([]);
    setStatus(null);
    setActiveStep(0);
    setIsProcessing(false);
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Paper elevation={3} sx={{ overflow: 'hidden' }}>
        {/* Header */}
        <GradientHeader>
          <Typography variant="h4" component="h1" gutterBottom>
            DOCX to PDF Converter
          </Typography>
          <Typography variant="subtitle1">
            Convert Word documents to PDF using Microsoft Word (WinAX)
          </Typography>
        </GradientHeader>

        {/* Stepper */}
        <Box sx={{ p: 3, backgroundColor: 'grey.50' }}>
          <Stepper activeStep={activeStep} alternativeLabel>
            {steps.map((label, index) => (
              <Step key={label}>
                <StepLabel
                  StepIconComponent={({ active, completed }) => {
                    const icons = [CloudUpload, Description, PictureAsPdf];
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
          {/* File Selection Step */}
          {activeStep === 0 && (
            <Fade in timeout={500}>
              <Box>
                <UploadArea onClick={selectFiles}>
                  <Description sx={{ fontSize: 64, color: 'grey.400', mb: 2 }} />
                  <Typography variant="h6" gutterBottom>
                    Select DOCX Files
                  </Typography>
                  <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                    Choose Word documents to convert to PDF format
                  </Typography>
                  <Button
                    variant="contained"
                    startIcon={<CloudUpload />}
                    size="large"
                    disabled={isProcessing}
                  >
                    {isProcessing ? 'Processing...' : 'Choose DOCX Files'}
                  </Button>
                </UploadArea>
                
                <Alert severity="info" sx={{ mt: 3 }}>
                  <Typography variant="body2">
                    💡 Requires Microsoft Word • Uses Windows COM automation (WinAX) • Supports .docx files<br/>
                    📁 PDFs will be automatically saved to your Downloads folder
                  </Typography>
                </Alert>
              </Box>
            </Fade>
          )}

          {/* Review Selection Step */}
          {activeStep === 1 && selectedFiles.length > 0 && (
            <Fade in timeout={500}>
              <Grid container spacing={3}>
                <Grid item xs={12} lg={8}>
                  <Card>
                    <CardContent>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
                        <Typography variant="h6">
                          Selected Files ({selectedFiles.length})
                        </Typography>
                        <Button 
                          size="small" 
                          onClick={selectFiles}
                          startIcon={<CloudUpload />}
                          disabled={isProcessing}
                        >
                          Add More Files
                        </Button>
                      </Box>
                      
                      <List>
                        {selectedFiles.map((file, index) => (
                          <ListItem
                            key={index}
                            sx={{
                              border: 1,
                              borderColor: 'grey.200',
                              borderRadius: 1,
                              mb: 1,
                              backgroundColor: 'grey.50'
                            }}
                          >
                            <ListItemIcon>
                              <Description color="primary" />
                            </ListItemIcon>
                            <ListItemText
                              primary={file.name}
                              secondary={
                                <Typography variant="caption" sx={{ fontFamily: 'monospace' }}>
                                  {file.path}
                                </Typography>
                              }
                            />
                            <IconButton
                              onClick={() => removeFile(index)}
                              disabled={isProcessing}
                              color="error"
                            >
                              <Delete />
                            </IconButton>
                          </ListItem>
                        ))}
                      </List>
                    </CardContent>
                  </Card>
                </Grid>

                <Grid item xs={12} lg={4}>
                  <Card sx={{ height: 'fit-content' }}>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        Output Location
                      </Typography>
                      
                      <Alert severity="success" sx={{ mb: 2 }}>
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                          <FolderOpen />
                          <Box>
                            <Typography variant="subtitle2">
                              Downloads Folder
                            </Typography>
                            <Typography variant="body2">
                              PDFs will be saved here automatically
                            </Typography>
                          </Box>
                        </Box>
                      </Alert>

                      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
                        Conversion Options
                      </Typography>
                      
                      <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                        <Button
                          fullWidth
                          variant="contained"
                          color="primary"
                          onClick={convertWithWinAX}
                          disabled={isProcessing}
                          startIcon={<Merge />}
                        >
                          Merge to Single PDF
                        </Button>

                        <Button
                          fullWidth
                          variant="contained"
                          color="success"
                          onClick={convertSeparately}
                          disabled={isProcessing}
                          startIcon={<FileCopy />}
                        >
                          Convert Separately
                        </Button>
                        
                        <Button
                          fullWidth
                          variant="outlined"
                          color="error"
                          onClick={clearAllFiles}
                          disabled={isProcessing}
                          startIcon={<Clear />}
                        >
                          Clear All
                        </Button>
                      </Box>
                    </CardContent>
                  </Card>
                </Grid>
              </Grid>
            </Fade>
          )}

          {/* Convert Step */}
          {activeStep === 2 && (
            <Fade in timeout={500}>
              <Box>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 3 }}>
                  <Typography variant="h6">
                    Conversion Progress
                  </Typography>
                  <Button 
                    onClick={() => setActiveStep(1)}
                    variant="outlined"
                    disabled={isProcessing}
                  >
                    Back to Review
                  </Button>
                </Box>

                <Card sx={{ mb: 3 }}>
                  <CardContent>
                    <Typography variant="subtitle1" gutterBottom>
                      Conversion Summary
                    </Typography>
                    <Grid container spacing={2}>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Files Selected:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {selectedFiles.length}
                        </Typography>
                      </Grid>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Output Location:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          Downloads folder
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
                    Convert More Files
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
            No
          </Button>
          <Button onClick={dialogContent.onConfirm} variant="contained" autoFocus>
            Yes
          </Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
};

export default DocToPdf;