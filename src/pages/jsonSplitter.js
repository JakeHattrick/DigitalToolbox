import React, { useState, useRef } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  Grid,
  Card,
  CardContent,
  TextField,
  FormControl,
  FormLabel,
  RadioGroup,
  FormControlLabel,
  Radio,
  Select,
  MenuItem,
  Checkbox,
  FormGroup,
  Alert,
  Divider,
  Chip,
  List,
  ListItem,
  ListItemText,
  Container,
  Fade,
  Collapse,
  LinearProgress
} from '@mui/material';
import {
  CloudUpload,
  Download,
  DataObject,
  ContentCut,
  Description,
  Folder,
  CheckCircle,
  Error as ErrorIcon
} from '@mui/icons-material';
import { styled } from '@mui/material/styles';

// Styled components
const GradientHeader = styled(Paper)(({ theme }) => ({
  background: 'linear-gradient(135deg, #27a7b0ff 0%, #25af43ff 100%)',
  color: 'white',
  padding: theme.spacing(4),
  borderRadius: 0,
  textAlign: 'center'
}));

const UploadArea = styled(Paper)(({ theme }) => ({
  border: `3px dashed ${theme.palette.primary.main}`,
  borderRadius: theme.shape.borderRadius * 2,
  padding: theme.spacing(8),
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.3s ease',
  position: 'relative',
  overflow: 'hidden',
  '&:hover': {
    borderColor: theme.palette.primary.dark,
    backgroundColor: theme.palette.action.hover,
    transform: 'scale(1.02)'
  },
  '&.dragover': {
    borderColor: theme.palette.primary.dark,
    backgroundColor: theme.palette.primary.light + '10',
    transform: 'scale(1.02)'
  }
}));

const OptionCard = styled(Card)(({ theme }) => ({
  marginBottom: theme.spacing(3),
  backgroundColor: theme.palette.grey[50]
}));

const ResultItem = styled(Paper)(({ theme }) => ({
  padding: theme.spacing(2),
  marginBottom: theme.spacing(2),
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  border: `1px solid ${theme.palette.divider}`
}));

function JsonSplitter() {
  const [jsonData, setJsonData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [splitMethod, setSplitMethod] = useState('chunks');
  const [chunkSize, setChunkSize] = useState(100);
  const [itemCount, setItemCount] = useState(1000);
  const [fileCount, setFileCount] = useState(2);
  const [baseFilename, setBaseFilename] = useState('');
  const [selectedColumn, setSelectedColumn] = useState('');
  const [keepAllColumns, setKeepAllColumns] = useState(true);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [availableColumns, setAvailableColumns] = useState([]);
  const [columnPreview, setColumnPreview] = useState('');
  const [splitFiles, setSplitFiles] = useState([]);
  const [showOptions, setShowOptions] = useState(false);
  const [showResults, setShowResults] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [alertMessage, setAlertMessage] = useState(null);
  const fileInputRef = useRef(null);

  const showAlert = (message, severity = 'info') => {
    setAlertMessage({ message, severity });
    setTimeout(() => setAlertMessage(null), 5000);
  };

  const handleFileUpload = async (file) => {
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.json')) {
      showAlert('Please select a valid JSON file.', 'error');
      return;
    }

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      setJsonData(parsed);
      setFileName(file.name.replace('.json', ''));
      setBaseFilename(file.name.replace('.json', ''));
      setShowOptions(true);
      setShowResults(false);
      showAlert(`File "${file.name}" loaded successfully!`, 'success');

      // Setup columns if it's an array
      if (Array.isArray(parsed) && parsed.length > 0) {
        const columns = extractColumns(parsed);
        setAvailableColumns(columns);
        setSelectedColumns(columns);
      }
    } catch (error) {
      showAlert('Invalid JSON file. Please check your file format.', 'error');
      setJsonData(null);
    }
  };

  const extractColumns = (data) => {
    const allKeys = new Set();
    const sampleSize = Math.min(100, data.length);
    
    for (let i = 0; i < sampleSize; i++) {
      const item = data[i];
      if (typeof item === 'object' && item !== null) {
        Object.keys(item).forEach(key => allKeys.add(key));
      }
    }
    
    return Array.from(allKeys).sort();
  };

  const handleColumnChange = (column) => {
    setSelectedColumn(column);
    if (column && jsonData) {
      const preview = getColumnPreview(column);
      setColumnPreview(preview);
    } else {
      setColumnPreview('');
    }
  };

  const getColumnPreview = (columnName) => {
    if (!Array.isArray(jsonData) || jsonData.length === 0) return '';
    
    const values = new Set();
    const sampleSize = Math.min(100, jsonData.length);
    
    for (let i = 0; i < sampleSize; i++) {
      const item = jsonData[i];
      if (typeof item === 'object' && item !== null && columnName in item) {
        values.add(String(item[columnName]));
      }
    }
    
    const uniqueValues = Array.from(values).sort();
    const preview = uniqueValues.slice(0, 10);
    
    if (uniqueValues.length === 0) {
      return 'No values found for this column.';
    } else {
      let previewText = `Found ${uniqueValues.length} unique value${uniqueValues.length !== 1 ? 's' : ''}: ${preview.join(', ')}`;
      if (uniqueValues.length > 10) {
        previewText += '...';
      }
      return previewText;
    }
  };

  const handleColumnSelection = (column, checked) => {
    if (checked) {
      setSelectedColumns([...selectedColumns, column]);
    } else {
      setSelectedColumns(selectedColumns.filter(c => c !== column));
      setKeepAllColumns(false);
    }
  };

  const handleKeepAllColumns = (checked) => {
    setKeepAllColumns(checked);
    if (checked) {
      setSelectedColumns(availableColumns);
    }
  };

  const splitJsonFile = () => {
    if (!jsonData) {
      showAlert('Please upload a JSON file first.', 'error');
      return;
    }

    const baseFile = baseFilename || 'split-data';
    const newSplitFiles = [];

    try {
      if (Array.isArray(jsonData)) {
        if (splitMethod === 'column') {
          splitByColumn(jsonData, baseFile, newSplitFiles);
        } else {
          splitArray(jsonData, splitMethod, baseFile, newSplitFiles);
        }
      } else if (typeof jsonData === 'object') {
        if (splitMethod === 'column') {
          showAlert('Column splitting is only available for JSON arrays.', 'error');
          return;
        }
        splitObject(jsonData, splitMethod, baseFile, newSplitFiles);
      } else {
        showAlert('JSON data must be an array or object to split.', 'error');
        return;
      }

      setSplitFiles(newSplitFiles);
      setShowResults(true);
      showAlert(`Successfully split into ${newSplitFiles.length} files!`, 'success');
    } catch (error) {
      showAlert('Error splitting file: ' + error.message, 'error');
    }
  };

  const splitArray = (data, method, baseFile, files) => {
    const columnsToKeep = keepAllColumns ? null : selectedColumns;
    let chunks = [];
    
    switch (method) {
      case 'chunks':
        for (let i = 0; i < data.length; i += chunkSize) {
          chunks.push(data.slice(i, i + chunkSize));
        }
        break;
        
      case 'count':
        for (let i = 0; i < data.length; i += itemCount) {
          chunks.push(data.slice(i, i + itemCount));
        }
        break;
        
      case 'files':
        const itemsPerFile = Math.ceil(data.length / fileCount);
        for (let i = 0; i < fileCount; i++) {
          const start = i * itemsPerFile;
          const end = Math.min(start + itemsPerFile, data.length);
          if (start < data.length) {
            chunks.push(data.slice(start, end));
          }
        }
        break;
    }

    chunks.forEach((chunk, index) => {
      const filteredChunk = columnsToKeep ? 
        chunk.map(item => filterObjectColumns(item, columnsToKeep)) : 
        chunk;
      
      const filename = `${baseFile}-part-${index + 1}.json`;
      const content = JSON.stringify(filteredChunk, null, 2);
      files.push({ filename, content, size: filteredChunk.length });
    });
  };

  const splitObject = (data, method, baseFile, files) => {
    const keys = Object.keys(data);
    let chunks = [];
    
    switch (method) {
      case 'chunks':
        for (let i = 0; i < keys.length; i += chunkSize) {
          const chunkKeys = keys.slice(i, i + chunkSize);
          const chunk = {};
          chunkKeys.forEach(key => chunk[key] = data[key]);
          chunks.push(chunk);
        }
        break;
        
      case 'count':
        for (let i = 0; i < keys.length; i += itemCount) {
          const chunkKeys = keys.slice(i, i + itemCount);
          const chunk = {};
          chunkKeys.forEach(key => chunk[key] = data[key]);
          chunks.push(chunk);
        }
        break;
        
      case 'files':
        const keysPerFile = Math.ceil(keys.length / fileCount);
        for (let i = 0; i < fileCount; i++) {
          const start = i * keysPerFile;
          const end = Math.min(start + keysPerFile, keys.length);
          if (start < keys.length) {
            const chunkKeys = keys.slice(start, end);
            const chunk = {};
            chunkKeys.forEach(key => chunk[key] = data[key]);
            chunks.push(chunk);
          }
        }
        break;
    }

    chunks.forEach((chunk, index) => {
      const filename = `${baseFile}-part-${index + 1}.json`;
      const content = JSON.stringify(chunk, null, 2);
      files.push({ filename, content, size: Object.keys(chunk).length });
    });
  };

  const splitByColumn = (data, baseFile, files) => {
    const columnName = selectedColumn;
    const columnsToKeep = keepAllColumns ? null : selectedColumns;
    
    if (!columnName) {
      showAlert('Please select a column to split by.', 'error');
      return;
    }
    
    const groups = {};
    
    data.forEach((item) => {
      if (typeof item === 'object' && item !== null) {
        const columnValue = item[columnName];
        const key = columnValue !== undefined && columnValue !== null ? String(columnValue) : 'null_or_undefined';
        
        if (!groups[key]) {
          groups[key] = [];
        }
        
        const filteredItem = filterObjectColumns(item, columnsToKeep);
        groups[key].push(filteredItem);
      } else {
        if (!groups['non_object_items']) {
          groups['non_object_items'] = [];
        }
        groups['non_object_items'].push(item);
      }
    });
    
    Object.keys(groups).sort().forEach(groupKey => {
      const groupData = groups[groupKey];
      const safeGroupKey = groupKey.replace(/[^a-zA-Z0-9-_]/g, '_');
      const filename = `${baseFile}-${columnName}-${safeGroupKey}.json`;
      const content = JSON.stringify(groupData, null, 2);
      files.push({ filename, content, size: groupData.length });
    });
    
    if (Object.keys(groups).length === 0) {
      showAlert('No groups found. The selected column may not exist in the data.', 'error');
    }
  };

  const filterObjectColumns = (obj, columnsToKeep) => {
    if (!columnsToKeep || !Array.isArray(columnsToKeep)) {
      return obj;
    }
    
    const filtered = {};
    columnsToKeep.forEach(column => {
      if (column in obj) {
        filtered[column] = obj[column];
      }
    });
    
    return filtered;
  };

  const downloadFile = (index) => {
    const file = splitFiles[index];
    const blob = new Blob([file.content], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = file.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const formatSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileUpload(files[0]);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setDragOver(false);
  };

  return (
    <Container maxWidth="lg" sx={{ py: 4 }}>
      <Paper elevation={3} sx={{ overflow: 'hidden' }}>
        {/* Header */}
        <GradientHeader>
          <Typography variant="h4" component="h1" gutterBottom>
            JSON Splitter
          </Typography>
          <Typography variant="subtitle1">
            Split large JSON files into smaller, manageable chunks
          </Typography>
        </GradientHeader>

        <Box sx={{ p: 3 }}>
          {/* Alert Messages */}
          {alertMessage && (
            <Alert 
              severity={alertMessage.severity} 
              sx={{ mb: 3 }}
              icon={alertMessage.severity === 'success' ? <CheckCircle /> : <ErrorIcon />}
            >
              {alertMessage.message}
            </Alert>
          )}

          {/* Upload Area */}
          <UploadArea
            className={dragOver ? 'dragover' : ''}
            onClick={() => fileInputRef.current?.click()}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
          >
            <DataObject sx={{ fontSize: 48, color: 'primary.main', mb: 2 }} />
            <Typography variant="h6" gutterBottom>
              Click to select or drag & drop your JSON file
            </Typography>
            <Typography variant="body2" color="text.secondary">
              Supports .json files up to 100MB
            </Typography>
            <input
              ref={fileInputRef}
              type="file"
              accept=".json"
              onChange={(e) => handleFileUpload(e.target.files[0])}
              style={{ display: 'none' }}
            />
          </UploadArea>

          {/* Options */}
          <Collapse in={showOptions}>
            <Fade in={showOptions} timeout={500}>
              <OptionCard>
                <CardContent>
                  <Grid container spacing={3}>
                    {/* Split Method */}
                    <Grid item xs={12}>
                      <FormControl component="fieldset">
                        <FormLabel component="legend">Split Method:</FormLabel>
                        <RadioGroup
                          row
                          value={splitMethod}
                          onChange={(e) => setSplitMethod(e.target.value)}
                        >
                          <FormControlLabel value="chunks" control={<Radio />} label="By Chunk Size" />
                          <FormControlLabel value="count" control={<Radio />} label="By Item Count" />
                          <FormControlLabel value="files" control={<Radio />} label="Into N Files" />
                          <FormControlLabel value="column" control={<Radio />} label="By Column Value" />
                        </RadioGroup>
                      </FormControl>
                    </Grid>

                    {/* Method-specific options */}
                    {splitMethod === 'chunks' && (
                      <Grid item xs={12} md={6}>
                        <TextField
                          fullWidth
                          type="number"
                          label="Items per chunk"
                          value={chunkSize}
                          onChange={(e) => setChunkSize(parseInt(e.target.value) || 1)}
                          inputProps={{ min: 1 }}
                        />
                      </Grid>
                    )}

                    {splitMethod === 'count' && (
                      <Grid item xs={12} md={6}>
                        <TextField
                          fullWidth
                          type="number"
                          label="Maximum items per file"
                          value={itemCount}
                          onChange={(e) => setItemCount(parseInt(e.target.value) || 1)}
                          inputProps={{ min: 1 }}
                        />
                      </Grid>
                    )}

                    {splitMethod === 'files' && (
                      <Grid item xs={12} md={6}>
                        <TextField
                          fullWidth
                          type="number"
                          label="Number of files"
                          value={fileCount}
                          onChange={(e) => setFileCount(parseInt(e.target.value) || 2)}
                          inputProps={{ min: 2 }}
                        />
                      </Grid>
                    )}

                    {splitMethod === 'column' && (
                      <Grid item xs={12}>
                        <FormControl fullWidth>
                          <Select
                            value={selectedColumn}
                            onChange={(e) => handleColumnChange(e.target.value)}
                            displayEmpty
                          >
                            <MenuItem value="">Select a column...</MenuItem>
                            {availableColumns.map((column) => (
                              <MenuItem key={column} value={column}>
                                {column}
                              </MenuItem>
                            ))}
                          </Select>
                        </FormControl>
                        {columnPreview && (
                          <Alert severity="info" sx={{ mt: 2 }}>
                            <Typography variant="body2">{columnPreview}</Typography>
                          </Alert>
                        )}
                      </Grid>
                    )}

                    {/* Base Filename */}
                    <Grid item xs={12} md={6}>
                      <TextField
                        fullWidth
                        label="Base filename (optional)"
                        placeholder="my-data"
                        value={baseFilename}
                        onChange={(e) => setBaseFilename(e.target.value)}
                      />
                    </Grid>

                    {/* Column Filter */}
                    {Array.isArray(jsonData) && jsonData.length > 0 && (
                      <Grid item xs={12}>
                        <Typography variant="subtitle2" gutterBottom>
                          Columns to keep in output files:
                        </Typography>
                        <FormControlLabel
                          control={
                            <Checkbox
                              checked={keepAllColumns}
                              onChange={(e) => handleKeepAllColumns(e.target.checked)}
                            />
                          }
                          label="Keep all columns"
                        />
                        <Collapse in={!keepAllColumns}>
                          <Paper sx={{ p: 2, maxHeight: 200, overflow: 'auto' }}>
                            <FormGroup>
                              {availableColumns.map((column) => (
                                <FormControlLabel
                                  key={column}
                                  control={
                                    <Checkbox
                                      checked={selectedColumns.includes(column)}
                                      onChange={(e) => handleColumnSelection(column, e.target.checked)}
                                    />
                                  }
                                  label={column}
                                />
                              ))}
                            </FormGroup>
                          </Paper>
                        </Collapse>
                      </Grid>
                    )}

                    {/* Split Button */}
                    <Grid item xs={12}>
                      <Button
                        fullWidth
                        variant="contained"
                        size="large"
                        onClick={splitJsonFile}
                        startIcon={<ContentCut />}
                        sx={{
                          background: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
                          '&:hover': {
                            transform: 'translateY(-2px)',
                            boxShadow: '0 10px 20px rgba(79, 172, 254, 0.3)'
                          }
                        }}
                      >
                        Split JSON File
                      </Button>
                    </Grid>
                  </Grid>
                </CardContent>
              </OptionCard>
            </Fade>
          </Collapse>

          {/* Results */}
          <Collapse in={showResults}>
            <Fade in={showResults} timeout={500}>
              <Box>
                <Typography variant="h6" gutterBottom sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                  <Folder />
                  Generated Files:
                </Typography>
                {splitFiles.map((file, index) => (
                  <ResultItem key={index} elevation={1}>
                    <Box>
                      <Typography variant="subtitle1" fontWeight="bold">
                        {file.filename}
                      </Typography>
                      <Typography variant="body2" color="text.secondary">
                        {file.size} items • {formatSize(file.content.length)}
                      </Typography>
                    </Box>
                    <Button
                      variant="contained"
                      startIcon={<Download />}
                      onClick={() => downloadFile(index)}
                    >
                      Download
                    </Button>
                  </ResultItem>
                ))}
              </Box>
            </Fade>
          </Collapse>
        </Box>
      </Paper>
    </Container>
  );
}

export default JsonSplitter;