import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

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
    <div className="min-h-screen bg-gradient-to-br from-blue-500 to-purple-600 p-5">
      <div className="max-w-6xl mx-auto bg-white bg-opacity-95 backdrop-blur-md rounded-3xl shadow-2xl overflow-hidden">
        <div className="bg-gradient-to-r from-blue-500 to-purple-600 text-white p-8 text-center">
          <h1 className="text-4xl font-bold mb-3">Document Merge Tool</h1>
          <p className="text-lg opacity-90">
            Merge Excel data with Word document templates seamlessly. Use merge fields like {`{{fieldname}}`} in your Word documents.
          </p>
        </div>
        
        <div className="p-10 space-y-8">
          {/* Reference Files Section */}
          <div className="bg-white rounded-2xl p-8 shadow-lg border border-gray-100">
            <h2 className="text-2xl font-semibold text-gray-800 mb-6 flex items-center gap-3">
              📄 Reference Files
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block w-full">
                  <input
                    type="file"
                    accept=".docx"
                    onChange={(e) => handleFileSelect(e, 'refTemplate')}
                    className="hidden"
                  />
                  <div className="flex items-center justify-center p-4 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl cursor-pointer hover:shadow-lg transition-all duration-300 hover:-translate-y-1 min-h-16 font-medium">
                    Choose Reference Template (.docx)
                  </div>
                </label>
                <div className={`mt-3 p-3 rounded-lg text-sm text-center ${
                  getFileStatus('refTemplate').className === 'file-uploaded' 
                    ? 'bg-green-100 text-green-800 border border-green-200' 
                    : 'bg-gray-100 text-gray-600 border border-gray-200'
                }`}>
                  {getFileStatus('refTemplate').text}
                </div>
              </div>
              <div>
                <label className="block w-full">
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(e) => handleFileSelect(e, 'refData')}
                    className="hidden"
                  />
                  <div className="flex items-center justify-center p-4 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl cursor-pointer hover:shadow-lg transition-all duration-300 hover:-translate-y-1 min-h-16 font-medium">
                    Choose Reference Data (.xlsx)
                  </div>
                </label>
                <div className={`mt-3 p-3 rounded-lg text-sm text-center ${
                  getFileStatus('refData').className === 'file-uploaded' 
                    ? 'bg-green-100 text-green-800 border border-green-200' 
                    : 'bg-gray-100 text-gray-600 border border-gray-200'
                }`}>
                  {getFileStatus('refData').text}
                </div>
              </div>
            </div>
          </div>

          {/* Document Files Section */}
          <div className="bg-white rounded-2xl p-8 shadow-lg border border-gray-100">
            <h2 className="text-2xl font-semibold text-gray-800 mb-6 flex items-center gap-3">
              📊 Document Files
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block w-full">
                  <input
                    type="file"
                    accept=".docx"
                    onChange={(e) => handleFileSelect(e, 'docTemplate')}
                    className="hidden"
                  />
                  <div className="flex items-center justify-center p-4 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl cursor-pointer hover:shadow-lg transition-all duration-300 hover:-translate-y-1 min-h-16 font-medium">
                    Choose Document Template (.docx)
                  </div>
                </label>
                <div className={`mt-3 p-3 rounded-lg text-sm text-center ${
                  getFileStatus('docTemplate').className === 'file-uploaded' 
                    ? 'bg-green-100 text-green-800 border border-green-200' 
                    : 'bg-gray-100 text-gray-600 border border-gray-200'
                }`}>
                  {getFileStatus('docTemplate').text}
                </div>
              </div>
              <div>
                <label className="block w-full">
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(e) => handleFileSelect(e, 'docData')}
                    className="hidden"
                  />
                  <div className="flex items-center justify-center p-4 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-xl cursor-pointer hover:shadow-lg transition-all duration-300 hover:-translate-y-1 min-h-16 font-medium">
                    Choose Document Data (.xlsx)
                  </div>
                </label>
                <div className={`mt-3 p-3 rounded-lg text-sm text-center ${
                  getFileStatus('docData').className === 'file-uploaded' 
                    ? 'bg-green-100 text-green-800 border border-green-200' 
                    : 'bg-gray-100 text-gray-600 border border-gray-200'
                }`}>
                  {getFileStatus('docData').text}
                </div>
              </div>
            </div>
          </div>

          {/* Output Settings Section */}
          <div className="bg-white rounded-2xl p-8 shadow-lg border border-gray-100">
            <h2 className="text-2xl font-semibold text-gray-800 mb-6 flex items-center gap-3">
              ⚙️ Output Settings
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-end">
              <div className="relative">
                <label className="absolute -top-2 left-4 bg-white px-2 text-sm font-semibold text-blue-600">
                  Output Folder Name
                </label>
                <input
                  type="text"
                  value={outputLocation}
                  onChange={(e) => setOutputLocation(e.target.value)}
                  placeholder="Enter output folder name"
                  className="w-full p-4 border-2 border-gray-200 rounded-xl text-base transition-all duration-300 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100"
                />
              </div>
              <button
                onClick={startMerge}
                disabled={!canMerge()}
                className={`w-full p-4 rounded-xl text-lg font-semibold transition-all duration-300 ${
                  canMerge()
                    ? 'bg-gradient-to-r from-green-500 to-emerald-600 text-white hover:shadow-lg hover:-translate-y-1'
                    : 'bg-gray-400 text-gray-200 cursor-not-allowed'
                }`}
              >
                {isProcessing ? 'Processing...' : 'Start Merge Process'}
              </button>
            </div>
          </div>

          {/* Status Section */}
          {status && (
            <div className={`p-6 rounded-xl font-medium text-center ${
              status.type === 'success' 
                ? 'bg-green-100 text-green-800 border border-green-200'
                : status.type === 'error'
                ? 'bg-red-100 text-red-800 border border-red-200'
                : 'bg-blue-100 text-blue-800 border border-blue-200'
            }`}>
              {status.message}
              {status.progress !== null && (
                <div className="w-full h-1 bg-gray-200 rounded-full overflow-hidden mt-3">
                  <div 
                    className="h-full bg-gradient-to-r from-blue-500 to-purple-600 transition-all duration-300"
                    style={{ width: `${status.progress}%` }}
                  />
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default FileMerge;