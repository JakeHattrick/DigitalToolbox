import React, { useState, useRef, useEffect } from 'react';
import { CloudUpload, FileText, Trash2, X, Download } from 'lucide-react';

// Access Electron APIs through the secure preload script
const electronAPI = window.electronAPI;

const DocToPdf = () => {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [status, setStatus] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const removeFile = (index) => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  const clearAllFiles = () => {
    setSelectedFiles([]);
    setStatus(null);
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
      }
    } catch (error) {
      showStatus('Error selecting files: ' + error.message, 'error');
    }
  };

  useEffect(() => {
    // Subscribe to progress events from the worker
    const unsubscribe = electronAPI?.onDocxToPdfProgress?.((p) => {
      // p is an integer 0–100; worker emits 0–80 converting, 80–100 merging
      // keep the message short; you already show a nice progress bar
      console.log('progress',p);
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
        const openFolder = await electronAPI.showMessageBox({
          type: 'question',
          buttons: ['Yes', 'No'],
          defaultId: 0,
          title: 'Conversion Complete',
          message: 'Documents converted successfully!',
          detail: `PDF saved as: ${outputFileName}\n\nWould you like to open the Downloads folder?`
        });
        if (openFolder.response === 0) await electronAPI.openDownloadsFolder();
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
        
        // Ask user if they want to open the Downloads folder
        const openFolder = await electronAPI.showMessageBox({
          type: 'question',
          buttons: ['Yes', 'No'],
          defaultId: 0,
          title: 'Conversion Complete',
          message: 'Documents converted successfully!',
          detail: `PDFs saved to Downloads folder\n\nWould you like to open the Downloads folder?`
        });

        if (openFolder.response === 0) {
          await electronAPI.openDownloadsFolder();
        }
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

  return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-500 via-purple-600 to-purple-800 p-5">
      <div className="max-w-4xl mx-auto bg-white/95 backdrop-blur-lg rounded-3xl shadow-2xl overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-indigo-500 to-purple-600 text-white p-8 text-center">
          <h1 className="text-4xl font-bold mb-2">DOCX to PDF Converter</h1>
          <p className="text-xl opacity-90">Convert Word documents to PDF using Microsoft Word (WinAX)</p>
        </div>

        <div className="p-8">
          {/* File Selection Section */}
          <div className="bg-white rounded-2xl shadow-lg border border-gray-100 mb-8">
            <div className="p-6">
              <h2 className="text-2xl font-semibold mb-4 flex items-center gap-2">
                <CloudUpload className="text-indigo-500" size={28} />
                Select DOCX Files
              </h2>
              
              <div className="text-center">
                <button
                  onClick={selectFiles}
                  disabled={isProcessing}
                  className={`px-8 py-4 rounded-xl font-semibold text-white text-lg transition-all duration-300 flex items-center gap-2 mx-auto ${
                    isProcessing
                      ? 'bg-gray-400 cursor-not-allowed'
                      : 'bg-gradient-to-r from-indigo-500 to-purple-600 hover:from-indigo-600 hover:to-purple-700 hover:transform hover:-translate-y-1 hover:shadow-lg'
                  }`}
                >
                  <FileText size={20} />
                  {isProcessing ? 'Processing...' : 'Choose DOCX Files'}
                </button>
                <p className="text-sm text-gray-500 mt-3">
                  PDFs will be saved to your Downloads folder
                </p>
              </div>
            </div>
          </div>

          {/* Output Location Display */}
          {selectedFiles.length > 0 && (
            <div className="bg-white rounded-2xl shadow-lg border border-gray-100 mb-8">
              <div className="p-6">
                <h2 className="text-2xl font-semibold mb-4 flex items-center gap-2">
                  Output Location
                </h2>
                
                <div className="p-4 bg-gradient-to-r from-green-50 to-green-100 rounded-xl border border-green-200">
                  <div className="flex items-center gap-2 text-green-800">
                    <div>
                      <div className="text-sm font-medium">PDFs will be saved to:</div>
                      <div className="font-medium text-green-900">
                        Downloads folder
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* File List Section */}
          {selectedFiles.length > 0 && (
            <div className="bg-white rounded-2xl shadow-lg border border-gray-100 mb-8">
              <div className="p-6">
                <h2 className="text-2xl font-semibold mb-4 flex items-center gap-2">
                  Selected Files ({selectedFiles.length})
                </h2>
                
                <div className="space-y-3">
                  {selectedFiles.map((file, index) => (
                    <div
                      key={index}
                      className="bg-white rounded-xl border-l-4 border-indigo-500 shadow-md p-4 flex items-center justify-between hover:shadow-lg transition-shadow duration-200"
                    >
                      <div className="flex-1">
                        <div className="flex items-center gap-2 font-medium text-gray-800">
                          <FileText className="text-indigo-500" size={20} />
                          {file.name}
                        </div>
                        <div className="text-xs text-gray-400 mt-1 truncate" title={file.path}>
                          {file.path}
                        </div>
                      </div>
                      <button
                        onClick={() => removeFile(index)}
                        disabled={isProcessing}
                        className="ml-4 p-2 text-red-500 hover:bg-red-50 rounded-lg transition-colors duration-200 disabled:opacity-50"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* Action Buttons Section */}
          <div className="bg-white rounded-2xl shadow-lg border border-gray-100 mb-8">
            <div className="p-6">
              <h2 className="text-2xl font-semibold mb-4 flex items-center gap-2">
                Conversion Options
              </h2>
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <button
                  onClick={convertWithWinAX}
                  disabled={selectedFiles.length === 0 || isProcessing}
                  className={`w-full py-4 px-6 rounded-xl font-semibold text-white text-lg transition-all duration-300 flex items-center justify-center gap-2 ${
                    selectedFiles.length === 0 || isProcessing
                      ? 'bg-gray-400 cursor-not-allowed'
                      : 'bg-gradient-to-r from-indigo-500 to-purple-600 hover:from-indigo-600 hover:to-purple-700 hover:transform hover:-translate-y-1 hover:shadow-lg'
                  }`}
                >
                  <Download size={20} />
                  {isProcessing ? 'Converting...' : 'Merge to Single PDF'}
                </button>

                <button
                  onClick={convertSeparately}
                  disabled={selectedFiles.length === 0 || isProcessing}
                  className={`w-full py-4 px-6 rounded-xl font-semibold text-white text-lg transition-all duration-300 flex items-center justify-center gap-2 ${
                    selectedFiles.length === 0 || isProcessing
                      ? 'bg-gray-400 cursor-not-allowed'
                      : 'bg-gradient-to-r from-green-500 to-green-600 hover:from-green-600 hover:to-green-700 hover:transform hover:-translate-y-1 hover:shadow-lg'
                  }`}
                >
                  <Download size={20} />
                  {isProcessing ? 'Converting...' : 'Convert Separately'}
                </button>
                
                <button
                  onClick={clearAllFiles}
                  disabled={selectedFiles.length === 0 || isProcessing}
                  className={`w-full py-4 px-6 rounded-xl font-semibold text-white text-lg transition-all duration-300 flex items-center justify-center gap-2 ${
                    selectedFiles.length === 0 || isProcessing
                      ? 'bg-gray-400 cursor-not-allowed'
                      : 'bg-gradient-to-r from-red-500 to-red-600 hover:from-red-600 hover:to-red-700 hover:transform hover:-translate-y-1 hover:shadow-lg'
                  }`}
                >
                  <X size={20} />
                  Clear All
                </button>
              </div>
            </div>
          </div>

          {/* Status Section */}
          {status && (
            <div
              className={`rounded-xl p-4 border ${
                status.type === 'success'
                  ? 'bg-green-50 border-green-200 text-green-800'
                  : status.type === 'error'
                  ? 'bg-red-50 border-red-200 text-red-800'
                  : 'bg-blue-50 border-blue-200 text-blue-800'
              }`}
            >
              <div className="font-medium text-lg mb-2">
                {status.message}
              </div>
              {status.progress !== null && (
                <div className="w-full bg-gray-200 rounded-full h-3 overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-indigo-500 to-purple-600 transition-all duration-300 ease-out rounded-full"
                    style={{ width: `${status.progress}%` }}
                  />
                </div>
              )}
            </div>
          )}

          {/* Footer */}
          <div className="text-center text-gray-500 text-sm mt-8">
            <p>Requires Microsoft Word • Uses Windows COM automation (WinAX) • Supports .docx files</p>
            <p className="mt-1">PDFs are automatically saved to your Downloads folder</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default DocToPdf;