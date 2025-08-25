import React, { useState, useRef } from 'react';
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
  TextField,
  Divider,
  Chip,
  Alert,
  LinearProgress,
  IconButton,
  Collapse,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Container,
  Fade,
  Zoom
} from '@mui/material';
import {
  CloudUpload,
  Download,
  DataObject,
  TableChart,
  ExpandMore,
  ExpandLess,
  Visibility,
  Settings,
  Refresh,
  FolderOpen,
  Description
} from '@mui/icons-material';
import { styled } from '@mui/material/styles';

// Styled components
const GradientHeader = styled(Paper)(({ theme }) => ({
  background: 'linear-gradient(135deg, #5337d1ff 0%, #27a7b0ff 100%)',
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

const JsonTree = styled(Paper)(({ theme }) => ({
  maxHeight: 400,
  overflow: 'auto',
  backgroundColor: theme.palette.grey[50],
  padding: theme.spacing(1)
}));

const PreviewCode = styled('pre')(({ theme }) => ({
  fontSize: '0.75rem',
  fontFamily: 'monospace',
  overflow: 'auto',
  whiteSpace: 'pre-wrap',
  maxHeight: 200,
  margin: 0,
  padding: theme.spacing(1),
  backgroundColor: theme.palette.background.paper,
  border: `1px solid ${theme.palette.divider}`,
  borderRadius: theme.shape.borderRadius
}));

const CsvPreview = styled('pre')(({ theme }) => ({
  fontSize: '0.75rem',
  fontFamily: 'monospace',
  overflow: 'auto',
  whiteSpace: 'pre',
  maxHeight: 300,
  margin: 0,
  padding: theme.spacing(2),
  backgroundColor: theme.palette.background.paper,
  border: `1px solid ${theme.palette.divider}`,
  borderRadius: theme.shape.borderRadius
}));

function JsonToCsv() {
  const [jsonData, setJsonData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [activeStep, setActiveStep] = useState(0);
  const [selectedPath, setSelectedPath] = useState([]);
  const [csvData, setCsvData] = useState('');
  const [expandedPaths, setExpandedPaths] = useState(new Set());
  const fileInputRef = useRef(null);

  const steps = ['Upload JSON', 'Explore Structure', 'Convert & Download'];

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      setJsonData(parsed);
      setFileName(file.name.replace('.json', ''));
      setActiveStep(1);
      setExpandedPaths(new Set(['root']));
    } catch (error) {
      alert('Invalid JSON file. Please check the format and try again.');
    }
  };

  const getValueAtPath = (obj, path) => {
    if (path.length === 0) return obj;
    return path.reduce((current, key) => {
      if (current && typeof current === 'object' && key in current) {
        return current[key];
      }
      return null;
    }, obj);
  };

  const isArray = (value) => Array.isArray(value);
  const isObject = (value) => value && typeof value === 'object' && !Array.isArray(value);

  const toggleExpanded = (path) => {
    const pathKey = path.join('.');
    const newExpanded = new Set(expandedPaths);
    if (newExpanded.has(pathKey)) {
      newExpanded.delete(pathKey);
    } else {
      newExpanded.add(pathKey);
    }
    setExpandedPaths(newExpanded);
  };

  const renderJsonHierarchy = (obj, path = [], depth = 0) => {
    if (!obj || depth > 10) return null;

    const pathKey = path.join('.');
    const isExpanded = expandedPaths.has(pathKey);
    const isRoot = path.length === 0;
    const isSelected = selectedPath.join('.') === pathKey;

    if (isArray(obj)) {
      const hasObjects = obj.length > 0 && isObject(obj[0]);
      const canSelect = hasObjects;

      return (
        <Box key={pathKey} sx={{ ml: isRoot ? 0 : 2 }}>
          <ListItem
            button
            onClick={() => {
              if (canSelect) {
                setSelectedPath([...path]);
              }
              toggleExpanded(path);
            }}
            sx={{
              borderRadius: 1,
              backgroundColor: isSelected && canSelect ? 'primary.light' : 'transparent',
              '&:hover': { backgroundColor: 'action.hover' },
              borderLeft: isSelected && canSelect ? 4 : 0,
              borderLeftColor: 'primary.main'
            }}
          >
            <ListItemIcon sx={{ minWidth: 32 }}>
              {isExpanded ? <ExpandLess /> : <ExpandMore />}
            </ListItemIcon>
            <ListItemText
              primary={
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                  <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                    [{obj.length}] {path[path.length - 1] || 'root'}
                  </Typography>
                  {canSelect && (
                    <Chip 
                      label="Can convert to CSV" 
                      size="small" 
                      color="success" 
                      variant="outlined"
                    />
                  )}
                </Box>
              }
            />
          </ListItem>
          <Collapse in={isExpanded}>
            {obj.length > 0 && renderJsonHierarchy(obj[0], [...path, '0'], depth + 1)}
          </Collapse>
        </Box>
      );
    }

    if (isObject(obj)) {
      const keys = Object.keys(obj);
      const canSelectAsRow = isRoot && keys.length > 0;

      return (
        <Box key={pathKey}>
          <ListItem
            button
            onClick={() => {
              if (canSelectAsRow) {
                setSelectedPath([...path]);
              }
              toggleExpanded(path);
            }}
            sx={{
              ml: isRoot ? 0 : 2,
              borderRadius: 1,
              backgroundColor: isSelected && canSelectAsRow ? 'warning.light' : 'transparent',
              '&:hover': { backgroundColor: 'action.hover' },
              borderLeft: isSelected && canSelectAsRow ? 4 : 0,
              borderLeftColor: 'warning.main'
            }}
          >
            <ListItemIcon sx={{ minWidth: 32 }}>
              {isExpanded ? <ExpandLess /> : <ExpandMore />}
            </ListItemIcon>
            <ListItemText
              primary={
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                  <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                    {`{${keys.length}} ${path[path.length - 1] || 'root'}`}
                  </Typography>
                  {canSelectAsRow && (
                    <Chip 
                      label="Can convert as single row" 
                      size="small" 
                      color="warning" 
                      variant="outlined"
                    />
                  )}
                </Box>
              }
            />
          </ListItem>
          <Collapse in={isExpanded}>
            <Box sx={{ ml: 2 }}>
              {keys.slice(0, 20).map(key => (
                renderJsonHierarchy(obj[key], [...path, key], depth + 1)
              ))}
              {keys.length > 20 && (
                <Typography variant="caption" color="text.secondary" sx={{ ml: 2 }}>
                  ... and {keys.length - 20} more properties
                </Typography>
              )}
            </Box>
          </Collapse>
        </Box>
      );
    }

    return (
      <ListItem key={pathKey} sx={{ ml: 2 }}>
        <ListItemText
          primary={
            <Typography variant="body2" sx={{ fontFamily: 'monospace', color: 'text.secondary' }}>
              {path[path.length - 1]}: {typeof obj === 'string' ? `"${obj.slice(0, 50)}${obj.length > 50 ? '...' : ''}"` : String(obj)}
            </Typography>
          }
        />
      </ListItem>
    );
  };

  const flattenObject = (obj, prefix = '') => {
    const flattened = {};
    
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const newKey = prefix ? `${prefix}.${key}` : key;
        
        if (isObject(obj[key])) {
          Object.assign(flattened, flattenObject(obj[key], newKey));
        } else if (isArray(obj[key])) {
          flattened[newKey] = JSON.stringify(obj[key]);
        } else {
          flattened[newKey] = obj[key];
        }
      }
    }
    
    return flattened;
  };

  const convertToCsv = () => {
    const dataAtPath = getValueAtPath(jsonData, selectedPath);
    
    if (isObject(dataAtPath) && selectedPath.length === 0) {
      const flattened = flattenObject(dataAtPath);
      const headers = Object.keys(flattened).sort();
      
      const csvRows = [
        headers.join(','),
        headers.map(header => {
          const value = flattened[header];
          if (value === null || value === undefined) {
            return '';
          }
          const stringValue = String(value);
          if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
            return `"${stringValue.replace(/"/g, '""')}"`;
          }
          return stringValue;
        }).join(',')
      ];

      setCsvData(csvRows.join('\n'));
      setActiveStep(2);
      return;
    }
    
    if (!isArray(dataAtPath)) {
      alert('Selected path does not contain an array. Please select an array to convert to CSV.');
      return;
    }

    if (dataAtPath.length === 0) {
      alert('Selected array is empty.');
      return;
    }

    const flattenedRows = dataAtPath.map(item => {
      if (isObject(item)) {
        return flattenObject(item);
      } else {
        return { value: item };
      }
    });

    const allHeaders = new Set();
    flattenedRows.forEach(row => {
      Object.keys(row).forEach(key => allHeaders.add(key));
    });

    const headers = Array.from(allHeaders).sort();

    const csvRows = [
      headers.join(','),
      ...flattenedRows.map(row => 
        headers.map(header => {
          const value = row[header];
          if (value === null || value === undefined) {
            return '';
          }
          const stringValue = String(value);
          if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
            return `"${stringValue.replace(/"/g, '""')}"`;
          }
          return stringValue;
        }).join(',')
      )
    ];

    setCsvData(csvRows.join('\n'));
    setActiveStep(2);
  };

  const downloadCsv = () => {
    const blob = new Blob([csvData], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${fileName || 'converted'}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const reset = () => {
    setJsonData(null);
    setFileName('');
    setActiveStep(0);
    setSelectedPath([]);
    setCsvData('');
    setExpandedPaths(new Set());
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const previewData = () => {
    const dataAtPath = getValueAtPath(jsonData, selectedPath);
    
    if (isObject(dataAtPath) && selectedPath.length === 0) {
      return JSON.stringify(dataAtPath, null, 2);
    }
    
    if (isArray(dataAtPath) && dataAtPath.length > 0) {
      return JSON.stringify(dataAtPath.slice(0, 3), null, 2);
    }
    return 'No preview available';
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Paper elevation={3} sx={{ overflow: 'hidden' }}>
        {/* Header */}
        <GradientHeader>
          <Typography variant="h4" component="h1" gutterBottom>
            JSON to CSV Converter
          </Typography>
          <Typography variant="subtitle1">
            Upload JSON, explore hierarchy, and convert to CSV
          </Typography>
        </GradientHeader>

        {/* Stepper */}
        <Box sx={{ p: 3, backgroundColor: 'grey.50' }}>
          <Stepper activeStep={activeStep} alternativeLabel>
            {steps.map((label, index) => (
              <Step key={label}>
                <StepLabel
                  StepIconComponent={({ active, completed }) => {
                    const icons = [CloudUpload, Visibility, Settings];
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
          {/* Upload Step */}
          {activeStep === 0 && (
            <Fade in timeout={500}>
              <Box>
                <UploadArea onClick={() => fileInputRef.current?.click()}>
                  <DataObject sx={{ fontSize: 64, color: 'grey.400', mb: 2 }} />
                  <Typography variant="h6" gutterBottom>
                    Upload JSON File
                  </Typography>
                  <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
                    Select a JSON file to analyze its structure and convert to CSV
                  </Typography>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".json"
                    onChange={handleFileUpload}
                    style={{ display: 'none' }}
                  />
                  <Button
                    variant="contained"
                    startIcon={<CloudUpload />}
                    size="large"
                    component="span"
                  >
                    Choose JSON File
                  </Button>
                </UploadArea>
              </Box>
            </Fade>
          )}

          {/* Explore Step */}
          {activeStep === 1 && (
            <Fade in timeout={500}>
              <Grid container spacing={3}>
                <Grid item xs={12} lg={6}>
                  <Card>
                    <CardContent>
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
                        <Typography variant="h6">
                          JSON Structure
                        </Typography>
                        <Button 
                          size="small" 
                          onClick={reset}
                          startIcon={<Refresh />}
                        >
                          Upload Different File
                        </Button>
                      </Box>
                      <JsonTree>
                        <List dense>
                          {renderJsonHierarchy(jsonData)}
                        </List>
                      </JsonTree>
                      <Alert severity="info" sx={{ mt: 2 }}>
                        <Typography variant="body2">
                          💡 Click on arrays containing objects to select them for CSV conversion<br/>
                          📋 Arrays are shown with [count] and objects with {'{count}'}
                        </Typography>
                      </Alert>
                    </CardContent>
                  </Card>
                </Grid>

                <Grid item xs={12} lg={6}>
                  <Card sx={{ height: 'fit-content' }}>
                    <CardContent>
                      <Typography variant="h6" gutterBottom>
                        Selection Details
                      </Typography>
                      {selectedPath.length >= 0 && jsonData ? (
                        <Box sx={{ space: 2 }}>
                          <Alert severity="info" sx={{ mb: 2 }}>
                            <Typography variant="subtitle2" gutterBottom>
                              Selected Path
                            </Typography>
                            <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                              {selectedPath.length === 0 ? 'root' : selectedPath.join(' → ')}
                            </Typography>
                          </Alert>
                          
                          <Alert severity="success" sx={{ mb: 2 }}>
                            <Typography variant="subtitle2" gutterBottom>
                              Preview (first 3 items)
                            </Typography>
                            <PreviewCode>
                              {previewData()}
                            </PreviewCode>
                          </Alert>
                        </Box>
                      ) : (
                        <Box sx={{ textAlign: 'center', py: 6 }}>
                          <TableChart sx={{ fontSize: 48, color: 'grey.400', mb: 2 }} />
                          <Typography color="text.secondary">
                            Select an array from the JSON structure to convert to CSV
                          </Typography>
                        </Box>
                      )}
                    </CardContent>
                    {selectedPath.length >= 0 && jsonData && (
                      <CardActions>
                        <Button
                          fullWidth
                          variant="contained"
                          color="success"
                          onClick={convertToCsv}
                          startIcon={<TableChart />}
                        >
                          Convert to CSV
                        </Button>
                      </CardActions>
                    )}
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
                    CSV Output
                  </Typography>
                  <Box sx={{ gap: 1, display: 'flex' }}>
                    <Button 
                      onClick={() => setActiveStep(1)}
                      variant="outlined"
                    >
                      Back to Explorer
                    </Button>
                    <Button
                      onClick={downloadCsv}
                      variant="contained"
                      startIcon={<Download />}
                    >
                      Download CSV
                    </Button>
                  </Box>
                </Box>

                <Card sx={{ mb: 3 }}>
                  <CardContent>
                    <Typography variant="subtitle1" gutterBottom>
                      Conversion Summary
                    </Typography>
                    <Grid container spacing={2}>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Source Path:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {selectedPath.join(' → ') || 'root'}
                        </Typography>
                      </Grid>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Rows:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {csvData.split('\n').length - 1}
                        </Typography>
                      </Grid>
                      <Grid item xs={12} md={4}>
                        <Typography variant="body2" color="text.secondary">
                          Columns:
                        </Typography>
                        <Typography variant="body2" sx={{ fontFamily: 'monospace' }}>
                          {csvData.split('\n')[0]?.split(',').length || 0}
                        </Typography>
                      </Grid>
                    </Grid>
                  </CardContent>
                </Card>

                <Card>
                  <CardContent>
                    <Typography variant="subtitle1" gutterBottom>
                      CSV Preview
                    </Typography>
                    <CsvPreview>
                      {csvData.split('\n').slice(0, 50).join('\n')}
                      {csvData.split('\n').length > 50 && '\n... (truncated)'}
                    </CsvPreview>
                  </CardContent>
                </Card>

                <Box sx={{ textAlign: 'center', mt: 3 }}>
                  <Button 
                    onClick={reset}
                    variant="outlined"
                    startIcon={<Refresh />}
                  >
                    Convert Another File
                  </Button>
                </Box>
              </Box>
            </Fade>
          )}
        </Box>
      </Paper>
    </Container>
  );
}

export default JsonToCsv;