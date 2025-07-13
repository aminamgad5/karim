class ETAInvoiceExporter {
  constructor() {
    this.invoiceData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessing = false;
    
    this.initializeElements();
    this.attachEventListeners();
    this.checkCurrentPage();
    this.setupProgressListener();
  }
  
  initializeElements() {
    this.elements = {
      countInfo: document.getElementById('countInfo'),
      totalCountText: document.getElementById('totalCountText'),
      status: document.getElementById('status'),
      closeBtn: document.getElementById('closeBtn'),
      jsonBtn: document.getElementById('jsonBtn'),
      excelBtn: document.getElementById('excelBtn'),
      pdfBtn: document.getElementById('pdfBtn'),
      progressContainer: null, // Will be created dynamically
      progressBar: null,
      progressText: null,
      checkboxes: {
        date: document.getElementById('option-date'),
        id: document.getElementById('option-id'),
        sellerId: document.getElementById('option-seller-id'),
        sellerName: document.getElementById('option-seller-name'),
        buyerId: document.getElementById('option-buyer-id'),
        buyerName: document.getElementById('option-buyer-name'),
        uuid: document.getElementById('option-uuid'),
        type: document.getElementById('option-type'),
        separateSeller: document.getElementById('option-separate-seller'),
        separateBuyer: document.getElementById('option-separate-buyer'),
        downloadDetails: document.getElementById('option-download-details'),
        combineAll: document.getElementById('option-combine-all'),
        downloadAll: document.getElementById('option-download-all')
      }
    };
    
    this.createProgressElements();
  }
  
  createProgressElements() {
    // Create progress container
    this.elements.progressContainer = document.createElement('div');
    this.elements.progressContainer.className = 'progress-container';
    this.elements.progressContainer.style.cssText = `
      margin-top: 15px;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 8px;
      border: 1px solid #e9ecef;
      display: none;
    `;
    
    // Create progress bar
    this.elements.progressBar = document.createElement('div');
    this.elements.progressBar.className = 'progress-bar';
    this.elements.progressBar.style.cssText = `
      width: 100%;
      height: 20px;
      background: #e9ecef;
      border-radius: 10px;
      overflow: hidden;
      margin-bottom: 10px;
      position: relative;
    `;
    
    const progressFill = document.createElement('div');
    progressFill.className = 'progress-fill';
    progressFill.style.cssText = `
      height: 100%;
      background: linear-gradient(90deg, #28a745, #20c997);
      border-radius: 10px;
      width: 0%;
      transition: width 0.3s ease;
      position: relative;
    `;
    
    this.elements.progressBar.appendChild(progressFill);
    
    // Create progress text
    this.elements.progressText = document.createElement('div');
    this.elements.progressText.className = 'progress-text';
    this.elements.progressText.style.cssText = `
      text-align: center;
      font-size: 14px;
      color: #495057;
      font-weight: 500;
    `;
    
    // Assemble progress container
    this.elements.progressContainer.appendChild(this.elements.progressBar);
    this.elements.progressContainer.appendChild(this.elements.progressText);
    
    // Add to DOM
    const statusElement = this.elements.status;
    statusElement.parentNode.insertBefore(this.elements.progressContainer, statusElement.nextSibling);
  }
  
  attachEventListeners() {
    this.elements.closeBtn.addEventListener('click', () => window.close());
    this.elements.excelBtn.addEventListener('click', () => this.handleExport('excel'));
    this.elements.jsonBtn.addEventListener('click', () => this.handleExport('json'));
    this.elements.pdfBtn.addEventListener('click', () => this.handleExport('pdf'));
    
    // Add listener for download all checkbox
    this.elements.checkboxes.downloadAll.addEventListener('change', (e) => {
      if (e.target.checked) {
        this.showMultiPageWarning();
      }
    });
  }
  
  setupProgressListener() {
    // Listen for progress updates from content script
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
      if (message.action === 'progressUpdate') {
        this.updateProgress(message.progress);
      }
    });
  }
  
  showMultiPageWarning() {
    const warningText = `تحذير: سيتم تحميل جميع الصفحات (${this.totalPages} صفحة) وقد يستغرق وقتاً أطول.`;
    this.showStatus(warningText, 'loading');
    
    setTimeout(() => {
      this.elements.status.textContent = '';
      this.elements.status.className = 'status';
    }, 3000);
  }
  
  async checkCurrentPage() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      
      if (!tab.url.includes('invoicing.eta.gov.eg')) {
        this.showStatus('يرجى الانتقال إلى بوابة الفواتير الإلكترونية المصرية', 'error');
        this.disableButtons();
        return;
      }
      
      await this.loadInvoiceData();
    } catch (error) {
      this.showStatus('خطأ في فحص الصفحة الحالية', 'error');
      console.error('Error:', error);
    }
  }
  
  async loadInvoiceData() {
    try {
      this.showStatus('جاري تحميل بيانات الفواتير...', 'loading');
      
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'getInvoiceData' });
      
      if (!response || !response.success) {
        throw new Error('فشل في الحصول على بيانات الفواتير');
      }
      
      this.invoiceData = response.data.invoices || [];
      this.totalCount = response.data.totalCount || this.invoiceData.length;
      this.currentPage = response.data.currentPage || 1;
      this.totalPages = response.data.totalPages || 1;
      
      this.updateUI();
      this.showStatus('تم تحميل البيانات بنجاح', 'success');
      
    } catch (error) {
      this.showStatus('خطأ في تحميل البيانات: ' + error.message, 'error');
      console.error('Load error:', error);
    }
  }
  
  updateUI() {
    const currentPageCount = this.invoiceData.length;
    this.elements.countInfo.textContent = `الصفحة الحالية: ${currentPageCount} فاتورة | المجموع: ${this.totalCount} فاتورة`;
    this.elements.totalCountText.textContent = this.totalCount;
    
    // Update download all option text
    const downloadAllLabel = this.elements.checkboxes.downloadAll.parentElement.querySelector('label');
    if (downloadAllLabel) {
      downloadAllLabel.innerHTML = `تحميل جميع الصفحات - <span id="totalCountText">${this.totalCount}</span> فاتورة (${this.totalPages} صفحة)`;
    }
  }
  
  getSelectedOptions() {
    const options = {};
    Object.keys(this.elements.checkboxes).forEach(key => {
      options[key] = this.elements.checkboxes[key].checked;
    });
    return options;
  }
  
  async handleExport(format) {
    if (this.isProcessing) {
      this.showStatus('جاري المعالجة... يرجى الانتظار', 'loading');
      return;
    }
    
    const options = this.getSelectedOptions();
    
    if (!this.validateOptions(options)) {
      return;
    }
    
    this.isProcessing = true;
    this.disableButtons();
    
    try {
      if (options.downloadAll) {
        await this.exportAllPages(format, options);
      } else {
        await this.exportCurrentPage(format, options);
      }
    } catch (error) {
      this.showStatus('خطأ في التصدير: ' + error.message, 'error');
      console.error('Export error:', error);
    } finally {
      this.isProcessing = false;
      this.enableButtons();
      this.hideProgress();
    }
  }
  
  validateOptions(options) {
    const hasBasicField = options.date || options.id || options.uuid;
    if (!hasBasicField) {
      this.showStatus('يرجى اختيار حقل واحد على الأقل للتصدير', 'error');
      return false;
    }
    return true;
  }
  
  async exportCurrentPage(format, options) {
    this.showStatus('جاري تصدير الصفحة الحالية...', 'loading');
    
    let dataToExport = [...this.invoiceData];
    
    if (options.downloadDetails) {
      this.showStatus('جاري تحميل تفاصيل الفواتير...', 'loading');
      dataToExport = await this.loadInvoiceDetails(dataToExport);
    }
    
    await this.generateFile(dataToExport, format, options);
    this.showStatus(`تم تصدير ${dataToExport.length} فاتورة بنجاح!`, 'success');
  }
  
  async exportAllPages(format, options) {
    this.showProgress();
    this.showStatus('جاري تحميل جميع الصفحات...', 'loading');
    
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    
    // Send message to content script to process all pages
    const allData = await chrome.tabs.sendMessage(tab.id, { 
      action: 'getAllPagesData',
      options: { ...options, progressCallback: true }
    });
    
    if (!allData || !allData.success) {
      throw new Error('فشل في تحميل جميع الصفحات: ' + (allData?.error || 'خطأ غير معروف'));
    }
    
    let dataToExport = allData.data;
    
    if (options.downloadDetails && dataToExport.length > 0) {
      this.updateProgress({
        currentPage: this.totalPages,
        totalPages: this.totalPages,
        message: 'جاري تحميل تفاصيل جميع الفواتير...'
      });
      
      dataToExport = await this.loadInvoiceDetails(dataToExport);
    }
    
    this.updateProgress({
      currentPage: this.totalPages,
      totalPages: this.totalPages,
      message: 'جاري إنشاء الملف...'
    });
    
    await this.generateFile(dataToExport, format, options);
    this.showStatus(`تم تصدير ${dataToExport.length} فاتورة من جميع الصفحات بنجاح!`, 'success');
  }
  
  showProgress() {
    this.elements.progressContainer.style.display = 'block';
    this.updateProgress({ currentPage: 0, totalPages: this.totalPages, message: 'جاري البدء...' });
  }
  
  hideProgress() {
    this.elements.progressContainer.style.display = 'none';
  }
  
  updateProgress(progress) {
    if (!this.elements.progressContainer || this.elements.progressContainer.style.display === 'none') {
      return;
    }
    
    const percentage = progress.totalPages > 0 ? (progress.currentPage / progress.totalPages) * 100 : 0;
    
    const progressFill = this.elements.progressBar.querySelector('.progress-fill');
    if (progressFill) {
      progressFill.style.width = `${Math.min(percentage, 100)}%`;
    }
    
    if (this.elements.progressText) {
      this.elements.progressText.textContent = progress.message || `الصفحة ${progress.currentPage} من ${progress.totalPages}`;
    }
  }
  
  async loadInvoiceDetails(invoices) {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const detailedInvoices = [];
    
    for (let i = 0; i < invoices.length; i++) {
      const invoice = invoices[i];
      this.showStatus(`جاري تحميل تفاصيل الفاتورة ${i + 1} من ${invoices.length}...`, 'loading');
      
      try {
        const detailResponse = await chrome.tabs.sendMessage(tab.id, {
          action: 'getInvoiceDetails',
          invoiceId: invoice.uuid
        });
        
        if (detailResponse && detailResponse.success) {
          detailedInvoices.push({
            ...invoice,
            details: detailResponse.data
          });
        } else {
          detailedInvoices.push(invoice);
        }
      } catch (error) {
        console.warn(`Failed to load details for invoice ${invoice.uuid}:`, error);
        detailedInvoices.push(invoice);
      }
      
      // Add small delay to avoid overwhelming the server
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    
    return detailedInvoices;
  }
  
  async generateFile(data, format, options) {
    switch (format) {
      case 'excel':
        this.generateInteractiveExcelFile(data, options);
        break;
      case 'json':
        this.generateJSONFile(data, options);
        break;
      case 'pdf':
        this.showStatus('تصدير PDF غير متاح حاليًا', 'error');
        break;
    }
  }
  
  generateInteractiveExcelFile(data, options) {
    const wb = XLSX.utils.book_new();
    
    // Create main summary sheet with interactive view buttons
    this.createInteractiveSummarySheet(wb, data, options);
    
    // Create details sheets for each invoice
    if (options.downloadDetails) {
      this.createDetailsSheets(wb, data, options);
    }
    
    // Add statistics sheet
    this.createStatisticsSheet(wb, data, options);
    
    const timestamp = new Date().toISOString().split('T')[0].replace(/-/g, '');
    const pageInfo = options.downloadAll ? 'AllPages' : `Page${this.currentPage}`;
    const filename = `ETA_Invoices_${pageInfo}_${timestamp}.xlsx`;
    
    XLSX.writeFile(wb, filename);
  }
  
  createInteractiveSummarySheet(wb, data, options) {
    // Arabic headers matching exactly the ETA portal interface
    const headers = [
      'رقم المستند الإلكتروني',    // Electronic Document ID
      'الرقم الداخلي',            // Internal ID  
      'تاريخ الإصدار',            // Issue Date
      'نوع المستند',             // Document Type
      'إصدار المستند',           // Document Version
      'إجمالي الفاتورة',          // Total Amount
      'اسم المورد',              // Supplier Name
      'الرقم الضريبي للمورد',      // Supplier Tax ID
      'اسم العميل',              // Customer Name
      'الرقم الضريبي للعميل',      // Customer Tax ID
      'رقم الإرسال',             // Submission ID
      'الحالة',                 // Status
      'رقم الصفحة'               // Page Number (for multi-page exports)
    ];
    
    const rows = [headers];
    
    data.forEach((invoice, index) => {
      const row = [
        invoice.documentId || '',        // رقم المستند الإلكتروني
        invoice.internalId || '',        // الرقم الداخلي
        invoice.issueDate || '',         // تاريخ الإصدار
        invoice.documentType || '',      // نوع المستند
        invoice.documentVersion || '',   // إصدار المستند
        invoice.totalAmount || '',       // إجمالي الفاتورة
        invoice.supplierName || '',      // اسم المورد
        invoice.supplierTaxId || '',     // الرقم الضريبي للمورد
        invoice.receiverName || '',      // اسم العميل
        invoice.receiverTaxId || '',     // الرقم الضريبي للعميل
        invoice.submissionId || '',      // رقم الإرسال
        invoice.status || '',            // الحالة
        invoice.pageNumber || ''         // رقم الصفحة
      ];
      rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // Format the worksheet
    this.formatWorksheet(ws, headers, data.length);
    
    XLSX.utils.book_append_sheet(wb, ws, 'ملخص الفواتير');
  }
  
  createStatisticsSheet(wb, data, options) {
    const stats = this.calculateStatistics(data);
    
    const statsData = [
      ['إحصائيات الفواتير', ''],
      ['', ''],
      ['إجمالي عدد الفواتير', data.length],
      ['إجمالي قيمة الفواتير', stats.totalValue.toFixed(2) + ' EGP'],
      ['إجمالي ضريبة القيمة المضافة', stats.totalVAT.toFixed(2) + ' EGP'],
      ['متوسط قيمة الفاتورة', stats.averageValue.toFixed(2) + ' EGP'],
      ['', ''],
      ['إحصائيات حسب الحالة', ''],
      ...Object.entries(stats.statusCounts).map(([status, count]) => [status, count]),
      ['', ''],
      ['إحصائيات حسب النوع', ''],
      ...Object.entries(stats.typeCounts).map(([type, count]) => [type, count]),
      ['', ''],
      ['تاريخ التصدير', new Date().toLocaleString('ar-EG')],
      ['نوع التصدير', options.downloadAll ? 'جميع الصفحات' : 'الصفحة الحالية']
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(statsData);
    
    // Format statistics sheet
    ws['!cols'] = [{ wch: 25 }, { wch: 20 }];
    
    XLSX.utils.book_append_sheet(wb, ws, 'الإحصائيات');
  }
  
  calculateStatistics(data) {
    const stats = {
      totalValue: 0,
      totalVAT: 0,
      averageValue: 0,
      statusCounts: {},
      typeCounts: {}
    };
    
    data.forEach(invoice => {
      // Calculate totals
      const value = parseFloat(invoice.totalAmount?.replace(/,/g, '') || 0);
      const vat = parseFloat(invoice.vatAmount || 0);
      
      stats.totalValue += value;
      stats.totalVAT += vat;
      
      // Count statuses
      const status = invoice.status || 'غير محدد';
      stats.statusCounts[status] = (stats.statusCounts[status] || 0) + 1;
      
      // Count types
      const type = invoice.type || 'غير محدد';
      stats.typeCounts[type] = (stats.typeCounts[type] || 0) + 1;
    });
    
    stats.averageValue = data.length > 0 ? stats.totalValue / data.length : 0;
    
    return stats;
  }
  
  formatInteractiveWorksheet(ws, headers, dataLength) {
  formatWorksheet(ws, headers, dataLength) {
    // Set column widths
    const colWidths = [
      { wch: 25 },  // رقم المستند الإلكتروني
      { wch: 15 },  // الرقم الداخلي
      { wch: 18 },  // تاريخ الإصدار
      { wch: 15 },  // نوع المستند
      { wch: 15 },  // إصدار المستند
      { wch: 15 },  // إجمالي الفاتورة
      { wch: 25 },  // اسم المورد
      { wch: 20 },  // الرقم الضريبي للمورد
      { wch: 25 },  // اسم العميل
      { wch: 20 },  // الرقم الضريبي للعميل
      { wch: 15 },  // رقم الإرسال
      { wch: 12 },  // الحالة
      { wch: 10 }   // رقم الصفحة
    ];
    
    ws['!cols'] = colWidths;
    
    // Style the header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "366092" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
    }
  }
  
  createDetailsSheets(wb, data, options) {
    data.forEach((invoice, index) => {
      if (invoice.details && invoice.details.length > 0) {
        const headers = [
          'إسم الصنف',           // Item name
          'كود الوحدة',          // Unit code
          'إسم الوحدة',          // Unit name
          'الكمية',             // Quantity
          'السعر',              // Price
          'القيمة',             // Value
          'ضريبة القيمة المضافة', // VAT
          'إجمالي'              // Total
        ];
        
        const rows = [headers];
        
        // Add invoice header info
        rows.push([
          `فاتورة رقم: ${invoice.internalId || index + 1}`,
          `التاريخ: ${invoice.issueDate || ''}`,
          `المورد: ${invoice.supplierName || ''}`,
          `العميل: ${invoice.receiverName || ''}`,
          `الإجمالي: ${invoice.totalAmount || ''} EGP`,
          '', '', ''
        ]);
        
        rows.push(['', '', '', '', '', '', '', '']); // Empty row
        
        // Add detail items
        invoice.details.forEach(item => {
          rows.push([
            item.name || '',
            item.unitCode || '',
            item.unitName || '',
            item.quantity || '',
            item.price || '',
            item.value || '',
            item.tax || '',
            item.total || ''
          ]);
        });
        
        // Add totals row
        const totalValue = invoice.details.reduce((sum, item) => sum + (parseFloat(item.value) || 0), 0);
        const totalTax = invoice.details.reduce((sum, item) => sum + (parseFloat(item.tax) || 0), 0);
        const grandTotal = totalValue + totalTax;
        
        rows.push(['', '', '', '', 'الإجمالي:', totalValue.toFixed(2), totalTax.toFixed(2), grandTotal.toFixed(2)]);
        
        const ws = XLSX.utils.aoa_to_sheet(rows);
        
        // Format the details sheet
        this.formatDetailsWorksheet(ws, headers);
        
        const sheetName = `تفاصيل_فاتورة_${index + 1}`;
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
    });
  }
  
  formatDetailsWorksheet(ws, headers) {
    const colWidths = [
      { wch: 25 }, // Item name
      { wch: 12 }, // Unit code
      { wch: 20 }, // Unit name
      { wch: 10 }, // Quantity
      { wch: 12 }, // Price
      { wch: 12 }, // Value
      { wch: 15 }, // VAT
      { wch: 12 }  // Total
    ];
    
    ws['!cols'] = colWidths;
    
    // Style the header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "2196F3" } },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }
  }
  
  generateJSONFile(data, options) {
    const jsonData = {
      exportDate: new Date().toISOString(),
      totalCount: data.length,
      exportType: options.downloadAll ? 'all_pages' : 'current_page',
      totalPages: this.totalPages,
      currentPage: this.currentPage,
      options: options,
      statistics: this.calculateStatistics(data),
      invoices: data
    };
    
    const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    const timestamp = new Date().toISOString().split('T')[0];
    const pageInfo = options.downloadAll ? 'AllPages' : `Page${this.currentPage}`;
    a.download = `ETA_Invoices_${pageInfo}_${timestamp}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  
  showStatus(message, type = '') {
    this.elements.status.textContent = message;
    this.elements.status.className = `status ${type}`;
    
    if (type === 'loading') {
      this.elements.status.innerHTML = `${message} <span class="loading-spinner"></span>`;
    }
    
    if (type === 'success' || type === 'error') {
      setTimeout(() => {
        if (!this.isProcessing) {
          this.elements.status.textContent = '';
          this.elements.status.className = 'status';
        }
      }, 3000);
    }
  }
  
  disableButtons() {
    this.elements.excelBtn.disabled = true;
    this.elements.jsonBtn.disabled = true;
    this.elements.pdfBtn.disabled = true;
  }
  
  enableButtons() {
    this.elements.excelBtn.disabled = false;
    this.elements.jsonBtn.disabled = false;
    this.elements.pdfBtn.disabled = false;
  }
}

// Initialize the exporter when popup loads
document.addEventListener('DOMContentLoaded', () => {
  new ETAInvoiceExporter();
});