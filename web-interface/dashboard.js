// HYCU M365 Sizing Dashboard JavaScript
class M365Dashboard {
    constructor() {
        this.reportData = null;
        this.storageChart = null;
        this.growthChart = null;
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const uploadArea = document.getElementById('upload-area');
        const fileInput = document.getElementById('file-input');

        // Drag and drop functionality
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.processFile(files[0]);
            }
        });

        // File input change
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                this.processFile(file);
            }
        });

        // Click to upload
        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });

        // Browse button click
        const browseButton = document.getElementById('browse-button');
        if (browseButton) {
            browseButton.addEventListener('click', (e) => {
                e.stopPropagation(); // Prevent event bubbling to upload area
                fileInput.click();
            });
        }
    }

    async processFile(file) {
        if (!file.name.endsWith('.html')) {
            return;
        }

        this.showLoading();
        
        try {
            const text = await file.text();
            // Store the original HTML content for the detailed report view
            window.dashboard = this;
            this.originalHtmlContent = text;
            
            this.parseReportData(text);
            this.displayDashboard();
            
            // Clear the file input after successful processing
            const fileInput = document.getElementById('file-input');
            if (fileInput) {
                fileInput.value = '';
            }
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing the report file. Please ensure it\'s a valid HYCU M365 report.');
            
            // Clear the file input even on error
            const fileInput = document.getElementById('file-input');
            if (fileInput) {
                fileInput.value = '';
            }
        }
    }

    parseReportData(htmlContent) {
        // Parse the HTML content to extract data
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlContent, 'text/html');
        
        this.reportData = {
            tenantInfo: this.extractTenantInfo(doc),
            storageData: this.extractStorageData(doc),
            growthData: this.extractGrowthData(doc),
            costAnalysis: this.extractCostAnalysis(doc),
            teamsData: this.extractTeamsData(doc),
            sitesData: this.extractSitesData(doc),
            licensingData: this.extractLicensingData(doc),
            mailboxData: this.extractMailboxData(doc),
            top5Data: this.extractTop5Data(doc)
        };
    }

    extractTenantInfo(doc) {
        const tenantName = this.extractTextByLabel(doc, 'Tenant Name') || 'Unknown';
        const totalUsers = this.extractNumberByLabel(doc, 'Total Users') || 0;
        const activeUsers = this.extractNumberByLabel(doc, 'Active Users') || 0;
        const guestUsers = this.extractNumberByLabel(doc, 'Guest Users') || 0;

        return { tenantName, totalUsers, activeUsers, guestUsers };
    }

    extractStorageData(doc) {
        const exchangeSize = this.extractNumberByLabel(doc, 'Exchange Online') || 0;
        const oneDriveSize = this.extractNumberByLabel(doc, 'OneDrive for Business') || 0;
        const sharePointSize = this.extractNumberByLabel(doc, 'SharePoint Online') || 0;
        const totalSize = this.extractNumberByLabel(doc, 'Total Storage') || 0;

        return { exchangeSize, oneDriveSize, sharePointSize, totalSize };
    }

    extractGrowthData(doc) {
        const currentSize = this.extractNumberByLabel(doc, 'Total Storage') || 0;
        const projections = {};
        
        // Extract growth projections from table
        const tableRows = doc.querySelectorAll('table tbody tr');
        tableRows.forEach(row => {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 3) {
                const rate = parseInt(cells[0].textContent.replace('%', ''));
                const projectedSize = parseFloat(cells[1].textContent);
                projections[rate] = projectedSize;
            }
        });

        return { currentSize, projections };
    }

    extractSitesData(doc) {
        const oneDriveAccounts = this.extractNumberByLabel(doc, 'OneDrive Accounts') || 0;
        const sharePointSites = this.extractNumberByLabel(doc, 'SharePoint Sites') || 0;
        const teamsSites = this.extractNumberByLabel(doc, 'Teams Sites') || 0;
        const totalSites = this.extractNumberByLabel(doc, 'Total Sites') || 0;

        return { oneDriveAccounts, sharePointSites, teamsSites, totalSites };
    }

    extractLicensingData(doc) {
        const licensedUsers = this.extractNumberByLabel(doc, 'Total Licensed Users') || 0;
        const hycuEntitlement = this.extractNumberByLabel(doc, 'HYCU Entitlement (50 GB/user)') || 0;
        const currentUsage = this.extractNumberByLabel(doc, 'Current Usage') || 0;
        const additionalLicenses = this.extractNumberByLabel(doc, 'Additional Licenses Needed') || 0;

        return { licensedUsers, hycuEntitlement, currentUsage, additionalLicenses };
    }

    extractMailboxData(doc) {
        const totalMailboxes = this.extractNumberByLabel(doc, 'Total Mailboxes') || 0;
        const regularMailboxes = this.extractNumberByLabel(doc, 'Regular Mailboxes') || 0;
        const sharedMailboxes = this.extractNumberByLabel(doc, 'Shared Mailboxes') || 0;
        const resourceMailboxes = this.extractNumberByLabel(doc, 'Resource Mailboxes') || 0;
        const archiveMailboxes = this.extractNumberByLabel(doc, 'Archive Mailboxes') || 0;
        const archivePercentage = this.extractNumberByLabel(doc, 'Archive %') || 0;
        const sharedAllowance = this.extractNumberByLabel(doc, '20% Allowance') || 0;
        const excessShared = this.extractNumberByLabel(doc, 'Excess Shared') || 0;

        return { 
            totalMailboxes, 
            regularMailboxes, 
            sharedMailboxes, 
            resourceMailboxes,
            archiveMailboxes,
            archivePercentage,
            sharedAllowance,
            excessShared
        };
    }

    extractTop5Data(doc) {
        // Extract Top 5 data from tables
        const top5Mailboxes = this.extractTop5FromTable(doc, 'Top 5 Mailboxes');
        const top5OneDrive = this.extractTop5FromTable(doc, 'Top 5 OneDrive');
        const top5SharePoint = this.extractTop5FromTable(doc, 'Top 5 SharePoint');

        return {
            mailboxes: top5Mailboxes,
            oneDrive: top5OneDrive,
            sharePoint: top5SharePoint
        };
    }

    extractTop5FromTable(doc, sectionTitle) {
        const data = [];
        
        // Find the section by looking for the heading
        const headings = doc.querySelectorAll('h3');
        let targetSection = null;
        
        for (let heading of headings) {
            if (heading.textContent.includes(sectionTitle)) {
                targetSection = heading.closest('div');
                break;
            }
        }
        
        if (targetSection) {
            const table = targetSection.querySelector('table tbody');
            if (table) {
                const rows = table.querySelectorAll('tr');
                rows.forEach(row => {
                    const cells = row.querySelectorAll('td');
                    if (cells.length >= 2) {
                        const name = cells[0].textContent.trim();
                        const size = parseFloat(cells[1].textContent.replace(/[^\d.-]/g, ''));
                        if (!isNaN(size)) {
                            data.push({ name, size });
                        }
                    }
                });
            }
        }
        
        return data.slice(0, 5); // Limit to top 5
    }

    extractCostAnalysis(doc) {
        const monthlyStorage = this.extractNumberByLabel(doc, 'Monthly Storage Cost') || 0;
        const monthlyUser = this.extractNumberByLabel(doc, 'Monthly User Cost') || 0;
        const totalMonthly = this.extractNumberByLabel(doc, 'Total Monthly Cost') || 0;
        const annual = this.extractNumberByLabel(doc, 'Annual Cost') || 0;

        return { monthlyStorage, monthlyUser, totalMonthly, annual };
    }

    extractTeamsData(doc) {
        const totalTeams = this.extractNumberByLabel(doc, 'Teams') || 0;
        const totalGroups = this.extractNumberByLabel(doc, 'Groups') || 0;
        
        // Extract Teams cost information
        const costPerMessage = this.extractTextByLabel(doc, 'Cost per message/notification');
        const costPerMillion = this.extractTextByLabel(doc, 'Cost per million messages');
        
        return { 
            totalTeams, 
            totalGroups,
            costPerMessage: costPerMessage || '$0.00075',
            costPerMillion: costPerMillion || '$750'
        };
    }

    extractTextByLabel(doc, label) {
        const elements = doc.querySelectorAll('.metric-label');
        for (let element of elements) {
            if (element.textContent.includes(label)) {
                const valueElement = element.previousElementSibling;
                return valueElement ? valueElement.textContent.trim() : null;
            }
        }
        return null;
    }

    extractNumberByLabel(doc, label) {
        const text = this.extractTextByLabel(doc, label);
        if (text) {
            const number = parseFloat(text.replace(/[^\d.-]/g, ''));
            return isNaN(number) ? 0 : number;
        }
        return 0;
    }

    showLoading() {
        document.getElementById('upload-section').classList.add('hidden');
        document.getElementById('loading-section').classList.remove('hidden');
        document.getElementById('dashboard-section').classList.add('hidden');
    }

    displayDashboard() {
        document.getElementById('loading-section').classList.add('hidden');
        document.getElementById('dashboard-section').classList.remove('hidden');
        
        this.populateTenantInfo();
        this.populateStorageData();
        this.populateGrowthData();
        this.populateCostAnalysis();
        this.populateSitesData();
        this.populateTeamsData();
        this.populateLicensingData();
        this.populateMailboxData();
        this.createCharts();
    }

    populateTenantInfo() {
        const data = this.reportData.tenantInfo;
        document.getElementById('tenant-name').textContent = data.tenantName;
        document.getElementById('total-users').textContent = data.totalUsers.toLocaleString();
        document.getElementById('active-users').textContent = data.activeUsers.toLocaleString();
        document.getElementById('guest-users').textContent = data.guestUsers.toLocaleString();
    }

    populateStorageData() {
        const data = this.reportData.storageData;
        document.getElementById('exchange-size').textContent = data.exchangeSize.toFixed(1);
        document.getElementById('onedrive-size').textContent = data.oneDriveSize.toFixed(1);
        document.getElementById('sharepoint-size').textContent = data.sharePointSize.toFixed(1);
        document.getElementById('total-size').textContent = data.totalSize.toFixed(1);
    }

    populateGrowthData() {
        // Growth data will be displayed in the chart
    }

    populateSitesData() {
        const data = this.reportData.sitesData;
        document.getElementById('onedrive-accounts').textContent = data.oneDriveAccounts.toLocaleString();
        document.getElementById('sharepoint-sites').textContent = data.sharePointSites.toLocaleString();
        document.getElementById('teams-sites').textContent = data.teamsSites.toLocaleString();
        document.getElementById('total-sites').textContent = data.totalSites.toLocaleString();
    }

    populateTeamsData() {
        const data = this.reportData.teamsData;
        document.getElementById('total-teams').textContent = data.totalTeams.toLocaleString();
        document.getElementById('total-groups').textContent = data.totalGroups.toLocaleString();
        document.getElementById('teams-cost-per-message').textContent = data.costPerMessage;
        document.getElementById('teams-cost-per-million').textContent = data.costPerMillion;
    }

    populateLicensingData() {
        const data = this.reportData.licensingData;
        document.getElementById('licensed-users').textContent = data.licensedUsers.toLocaleString();
        document.getElementById('hycu-entitlement').textContent = data.hycuEntitlement.toLocaleString();
        document.getElementById('current-usage').textContent = data.currentUsage.toLocaleString();
        document.getElementById('additional-licenses').textContent = data.additionalLicenses.toLocaleString();
    }

    populateMailboxData() {
        const data = this.reportData.mailboxData;
        document.getElementById('total-mailboxes').textContent = data.totalMailboxes.toLocaleString();
        document.getElementById('regular-mailboxes').textContent = data.regularMailboxes.toLocaleString();
        document.getElementById('shared-mailboxes').textContent = data.sharedMailboxes.toLocaleString();
        document.getElementById('resource-mailboxes').textContent = data.resourceMailboxes.toLocaleString();
        document.getElementById('archive-mailboxes').textContent = data.archiveMailboxes.toLocaleString();
        document.getElementById('archive-percentage').textContent = data.archivePercentage.toFixed(1) + '%';
        document.getElementById('shared-allowance').textContent = data.sharedAllowance.toLocaleString();
        document.getElementById('excess-shared').textContent = data.excessShared.toLocaleString();
    }

    populateCostAnalysis() {
        const data = this.reportData.costAnalysis;
        document.getElementById('monthly-storage-cost').textContent = `$${data.monthlyStorage.toFixed(2)}`;
        document.getElementById('monthly-user-cost').textContent = `$${data.monthlyUser.toFixed(2)}`;
        document.getElementById('total-monthly-cost').textContent = `$${data.totalMonthly.toFixed(2)}`;
        document.getElementById('annual-cost').textContent = `$${data.annual.toFixed(2)}`;
    }

    createCharts() {
        this.createStorageDistributionChart();
        this.createTop5Charts();
    }



    createTop5Charts() {
        this.createTop5MailboxesChart();
        this.createTop5OneDriveChart();
        this.createTop5SharePointChart();
    }

    createTop5MailboxesChart() {
        const ctx = document.getElementById('top5-mailboxes-chart').getContext('2d');
        const data = this.reportData.top5Data.mailboxes;
        
        if (data.length === 0) {
            ctx.fillText('No data available', 10, 50);
            return;
        }

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: data.map(item => item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name),
                datasets: [{
                    label: 'Size (GB)',
                    data: data.map(item => item.size),
                    backgroundColor: '#667eea',
                    borderColor: '#764ba2',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Size (GB)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    }

    createTop5OneDriveChart() {
        const ctx = document.getElementById('top5-onedrive-chart').getContext('2d');
        const data = this.reportData.top5Data.oneDrive;
        
        if (data.length === 0) {
            ctx.fillText('No data available', 10, 50);
            return;
        }

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: data.map(item => item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name),
                datasets: [{
                    label: 'Size (GB)',
                    data: data.map(item => item.size),
                    backgroundColor: '#764ba2',
                    borderColor: '#667eea',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Size (GB)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    }

    createTop5SharePointChart() {
        const ctx = document.getElementById('top5-sharepoint-chart').getContext('2d');
        const data = this.reportData.top5Data.sharePoint;
        
        if (data.length === 0) {
            ctx.fillText('No data available', 10, 50);
            return;
        }

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: data.map(item => item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name),
                datasets: [{
                    label: 'Size (GB)',
                    data: data.map(item => item.size),
                    backgroundColor: '#f093fb',
                    borderColor: '#667eea',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Size (GB)'
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });
    }

    createStorageDistributionChart() {
        const ctx = document.getElementById('storage-distribution-chart').getContext('2d');
        const data = this.reportData.storageData;
        
        // Calculate percentages
        const total = data.exchangeSize + data.oneDriveSize + data.sharePointSize;
        const exchangePercent = total > 0 ? (data.exchangeSize / total * 100).toFixed(1) : 0;
        const oneDrivePercent = total > 0 ? (data.oneDriveSize / total * 100).toFixed(1) : 0;
        const sharePointPercent = total > 0 ? (data.sharePointSize / total * 100).toFixed(1) : 0;

        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: [
                    `Exchange Online (${exchangePercent}%)`,
                    `OneDrive for Business (${oneDrivePercent}%)`,
                    `SharePoint Online (${sharePointPercent}%)`
                ],
                datasets: [{
                    data: [data.exchangeSize, data.oneDriveSize, data.sharePointSize],
                    backgroundColor: [
                        '#667eea',
                        '#764ba2',
                        '#f093fb'
                    ],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 20,
                            usePointStyle: true
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.parsed;
                                const percentage = ((value / total) * 100).toFixed(1);
                                return `${label}: ${value.toFixed(1)} GB (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
    }
}

// Export functions
function exportToPDF() {
    if (!window.dashboard || !window.dashboard.reportData) {
        return;
    }

    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        
        // Set up fonts and colors
        doc.setFont('helvetica');
        const primaryColor = [102, 126, 234]; // #667eea
        const secondaryColor = [118, 75, 162]; // #764ba2
        const accentColor = [240, 147, 251]; // #f093fb
        const darkGray = [60, 60, 60];
        const lightGray = [120, 120, 120];
        
        // Header with gradient background
        doc.setFillColor(...primaryColor);
        doc.rect(0, 0, 210, 35, 'F');
        
        
        // Title
        doc.setFontSize(20);
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.text('HYCU M365 Sizing Report', 20, 18);
        
        // Subtitle
        doc.setFontSize(10);
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'normal');
        doc.text('Comprehensive Microsoft 365 Tenant Analysis for Backup Planning', 20, 23);
        
        // Date
        doc.setFontSize(9);
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'normal');
        doc.text(`Generated on: ${new Date().toLocaleDateString()}`, 20, 30);
        
        let yPosition = 40;
        
        // Tenant Overview
        if (window.dashboard.reportData.tenantInfo) {
            yPosition = addSectionToPDF(doc, 'Tenant Overview', window.dashboard.reportData.tenantInfo, yPosition);
        }
        
        // Storage Overview
        if (window.dashboard.reportData.storageData) {
            yPosition = addStorageSectionToPDF(doc, 'Storage Overview', window.dashboard.reportData.storageData, yPosition);
        }
        
        // HYCU License Sizer
        if (window.dashboard.reportData.licensingData) {
            yPosition = addSectionToPDF(doc, 'HYCU License Sizer', window.dashboard.reportData.licensingData, yPosition);
        }
        
        // Service Overviews
        yPosition = addServiceOverviewToPDF(doc, 'Service Overviews', window.dashboard.reportData, yPosition);
        
        // Top 5 Analysis
        if (window.dashboard.reportData.top5Data) {
            yPosition = addTop5ToPDF(doc, 'Top 5 by Size', window.dashboard.reportData.top5Data, yPosition);
        }
        
        // Cost Estimation
        if (window.dashboard.reportData.costAnalysis) {
            yPosition = addCostEstimationToPDF(doc, 'Cost Estimation', window.dashboard.reportData.costAnalysis, yPosition);
        }
        
        // Add a note if no data is available
        if (yPosition <= 60) {
            doc.setFontSize(12);
            doc.setTextColor(100, 100, 100);
            doc.text('No data available for export. Please upload a valid HYCU M365 report.', 20, yPosition);
        }
        
        // Save the PDF
        doc.save('HYCU-M365-Sizing-Report.pdf');
        
    } catch (error) {
        console.error('Error generating PDF:', error);
    }
}

function addSectionToPDF(doc, title, data, yPosition) {
    // Check if we need a new page
    if (yPosition > 250) {
        doc.addPage();
        yPosition = 20;
    }
    
    // Section title
    doc.setFontSize(14);
    doc.setTextColor(102, 126, 234);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 20, yPosition);
    yPosition += 8;
    
    // Add a line under the title
    doc.setDrawColor(102, 126, 234);
    doc.setLineWidth(0.5);
    doc.line(20, yPosition, 190, yPosition);
    yPosition += 8;
    
    // Add metrics in a simple list format
    if (data) {
        const metrics = Object.entries(data).slice(0, 6);
        
        metrics.forEach(([key, value]) => {
            if (value !== undefined && value !== null && yPosition < 280) {
                // Metric label
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                doc.text(`${label}:`, 25, yPosition);
                
                // Metric value
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                const displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                doc.text(displayValue, 120, yPosition);
                
                yPosition += 6;
            }
        });
    }
    
    return yPosition + 10;
}

function addStorageSectionToPDF(doc, title, data, yPosition) {
    // Check if we need a new page
    if (yPosition > 250) {
        doc.addPage();
        yPosition = 20;
    }
    
    // Section title
    doc.setFontSize(14);
    doc.setTextColor(102, 126, 234);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 20, yPosition);
    yPosition += 8;
    
    // Add a line under the title
    doc.setDrawColor(102, 126, 234);
    doc.setLineWidth(0.5);
    doc.line(20, yPosition, 190, yPosition);
    yPosition += 8;
    
    if (data) {
        const storageMetrics = [
            { label: 'Exchange (GB)', value: data.exchangeSize },
            { label: 'OneDrive (GB)', value: data.oneDriveSize },
            { label: 'SharePoint (GB)', value: data.sharePointSize },
            { label: 'Total (GB)', value: data.totalSize }
        ];
        
        storageMetrics.forEach(metric => {
            if (metric.value !== undefined && metric.value !== null && yPosition < 280) {
                // Metric label
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                doc.text(`${metric.label}:`, 25, yPosition);
                
                // Metric value
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                doc.text(metric.value.toLocaleString(), 120, yPosition);
                
                yPosition += 6;
            }
        });
    }
    
    return yPosition + 10;
}

function addServiceOverviewToPDF(doc, title, data, yPosition) {
    // Check if we need a new page
    if (yPosition > 250) {
        doc.addPage();
        yPosition = 20;
    }
    
    // Main section title
    doc.setFontSize(14);
    doc.setTextColor(102, 126, 234);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 20, yPosition);
    yPosition += 8;
    
    // Add a line under the title
    doc.setDrawColor(102, 126, 234);
    doc.setLineWidth(0.5);
    doc.line(20, yPosition, 190, yPosition);
    yPosition += 8;
    
    // Mailbox Overview
    if (data.mailboxData) {
        doc.setFontSize(12);
        doc.setTextColor(118, 75, 162);
        doc.setFont('helvetica', 'bold');
        doc.text('Mailbox Overview', 25, yPosition);
        yPosition += 6;
        
        const mailboxMetrics = Object.entries(data.mailboxData).slice(0, 4);
        mailboxMetrics.forEach(([key, value]) => {
            if (value !== undefined && value !== null && yPosition < 280) {
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                doc.text(`${label}:`, 30, yPosition);
                
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                const displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                doc.text(displayValue, 120, yPosition);
                
                yPosition += 5;
            }
        });
        yPosition += 5;
    }
    
    // Sites & OneDrive Overview
    if (data.sitesData) {
        doc.setFontSize(12);
        doc.setTextColor(118, 75, 162);
        doc.setFont('helvetica', 'bold');
        doc.text('Sites & OneDrive Overview', 25, yPosition);
        yPosition += 6;
        
        const sitesMetrics = Object.entries(data.sitesData).slice(0, 4);
        sitesMetrics.forEach(([key, value]) => {
            if (value !== undefined && value !== null && yPosition < 280) {
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                doc.text(`${label}:`, 30, yPosition);
                
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                const displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                doc.text(displayValue, 120, yPosition);
                
                yPosition += 5;
            }
        });
        yPosition += 5;
    }
    
    // Teams Overview
    if (data.teamsData) {
        doc.setFontSize(12);
        doc.setTextColor(118, 75, 162);
        doc.setFont('helvetica', 'bold');
        doc.text('Teams Overview', 25, yPosition);
        yPosition += 6;
        
        const teamsMetrics = Object.entries(data.teamsData).slice(0, 4);
        teamsMetrics.forEach(([key, value]) => {
            if (value !== undefined && value !== null && yPosition < 280) {
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                doc.text(`${label}:`, 30, yPosition);
                
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                const displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                doc.text(displayValue, 120, yPosition);
                
                yPosition += 5;
            }
        });
    }
    
    return yPosition + 10;
}

function addServiceCardToPDF(doc, title, data, x, y, width, height, color) {
    // Card background
    doc.setFillColor(248, 249, 250);
    doc.rect(x, y, width, height, 'F');
    
    // Card border
    doc.setDrawColor(...color);
    doc.setLineWidth(0.5);
    doc.rect(x, y, width, height);
    
    // Title
    doc.setFontSize(12);
    doc.setTextColor(...color);
    doc.setFont('helvetica', 'bold');
    doc.text(title, x + 3, y + 8);
    
    // Add metrics
    if (data) {
        const metrics = Object.entries(data).slice(0, 4);
        let currentY = y + 15;
        
        metrics.forEach(([key, value]) => {
            if (value !== undefined && value !== null && currentY < y + height - 5) {
                // Metric background
                doc.setFillColor(255, 255, 255);
                doc.rect(x + 2, currentY - 1, width - 4, 6, 'F');
                
                // Metric value
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                const displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                doc.text(displayValue, x + 4, currentY + 2);
                
                // Metric label
                doc.setFontSize(7);
                doc.setTextColor(100, 100, 100);
                doc.setFont('helvetica', 'normal');
                const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
                doc.text(label, x + 4, currentY + 4);
                
                currentY += 8;
            }
        });
    }
}

function addTop5ToPDF(doc, title, data, yPosition) {
    // Check if we need a new page
    if (yPosition > 250) {
        doc.addPage();
        yPosition = 20;
    }
    
    // Section title
    doc.setFontSize(14);
    doc.setTextColor(102, 126, 234);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 20, yPosition);
    yPosition += 8;
    
    // Add a line under the title
    doc.setDrawColor(102, 126, 234);
    doc.setLineWidth(0.5);
    doc.line(20, yPosition, 190, yPosition);
    yPosition += 8;
    
    if (data) {
        // Top 5 Mailboxes
        if (data.mailboxes && data.mailboxes.length > 0) {
            doc.setFontSize(12);
            doc.setTextColor(118, 75, 162);
            doc.setFont('helvetica', 'bold');
            doc.text('Top 5 Mailboxes', 25, yPosition);
            yPosition += 6;
            
            data.mailboxes.slice(0, 5).forEach((item, index) => {
                if (yPosition < 280) {
                    doc.setFontSize(10);
                    doc.setTextColor(0, 0, 0);
                    doc.setFont('helvetica', 'normal');
                    doc.text(`${index + 1}. ${item.name}: ${item.size} GB`, 30, yPosition);
                    yPosition += 5;
                }
            });
            yPosition += 5;
        }
        
        // Top 5 OneDrive
        if (data.oneDrive && data.oneDrive.length > 0) {
            doc.setFontSize(12);
            doc.setTextColor(118, 75, 162);
            doc.setFont('helvetica', 'bold');
            doc.text('Top 5 OneDrive', 25, yPosition);
            yPosition += 6;
            
            data.oneDrive.slice(0, 5).forEach((item, index) => {
                if (yPosition < 280) {
                    doc.setFontSize(10);
                    doc.setTextColor(0, 0, 0);
                    doc.setFont('helvetica', 'normal');
                    doc.text(`${index + 1}. ${item.name}: ${item.size} GB`, 30, yPosition);
                    yPosition += 5;
                }
            });
            yPosition += 5;
        }
        
        // Top 5 SharePoint
        if (data.sharePoint && data.sharePoint.length > 0) {
            doc.setFontSize(12);
            doc.setTextColor(118, 75, 162);
            doc.setFont('helvetica', 'bold');
            doc.text('Top 5 SharePoint', 25, yPosition);
            yPosition += 6;
            
            data.sharePoint.slice(0, 5).forEach((item, index) => {
                if (yPosition < 280) {
                    doc.setFontSize(10);
                    doc.setTextColor(0, 0, 0);
                    doc.setFont('helvetica', 'normal');
                    doc.text(`${index + 1}. ${item.name}: ${item.size} GB`, 30, yPosition);
                    yPosition += 5;
                }
            });
        }
    }
    
    return yPosition + 10;
}

function addTop5CardToPDF(doc, title, data, x, y, width, height, color) {
    // Card background
    doc.setFillColor(248, 249, 250);
    doc.rect(x, y, width, height, 'F');
    
    // Card border
    doc.setDrawColor(...color);
    doc.setLineWidth(0.5);
    doc.rect(x, y, width, height);
    
    // Title
    doc.setFontSize(12);
    doc.setTextColor(...color);
    doc.setFont('helvetica', 'bold');
    doc.text(title, x + 3, y + 8);
    
    // Add top 5 items
    if (data && data.length > 0) {
        let currentY = y + 15;
        
        data.slice(0, 5).forEach((item, index) => {
            if (currentY < y + height - 5) {
                // Item background
                doc.setFillColor(255, 255, 255);
                doc.rect(x + 2, currentY - 1, width - 4, 6, 'F');
                
                // Rank
                doc.setFontSize(8);
                doc.setTextColor(...color);
                doc.setFont('helvetica', 'bold');
                doc.text(`${index + 1}.`, x + 4, currentY + 2);
                
                // Name (truncated)
                const name = item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name;
                doc.setFontSize(7);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                doc.text(name, x + 8, currentY + 2);
                
                // Size
                doc.setFontSize(7);
                doc.setTextColor(100, 100, 100);
                doc.text(`${item.size} GB`, x + 4, currentY + 4);
                
                currentY += 8;
            }
        });
    }
}

function addCostEstimationToPDF(doc, title, data, yPosition) {
    // Check if we need a new page
    if (yPosition > 250) {
        doc.addPage();
        yPosition = 20;
    }
    
    // Section title
    doc.setFontSize(14);
    doc.setTextColor(102, 126, 234);
    doc.setFont('helvetica', 'bold');
    doc.text(title, 20, yPosition);
    yPosition += 8;
    
    // Add a line under the title
    doc.setDrawColor(102, 126, 234);
    doc.setLineWidth(0.5);
    doc.line(20, yPosition, 190, yPosition);
    yPosition += 8;
    
    if (data) {
        const costMetrics = [
            { label: 'Monthly Storage Cost', value: data.monthlyStorage },
            { label: 'Monthly User Cost', value: data.monthlyUser },
            { label: 'Total Monthly Cost', value: data.totalMonthly },
            { label: 'Annual Cost', value: data.annual }
        ];
        
        costMetrics.forEach(metric => {
            if (metric.value !== undefined && metric.value !== null && yPosition < 280) {
                // Metric label
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'normal');
                doc.text(`${metric.label}:`, 25, yPosition);
                
                // Metric value
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                doc.setFont('helvetica', 'bold');
                doc.text(`$${metric.value.toFixed(2)}`, 120, yPosition);
                
                yPosition += 6;
            }
        });
        
        // Add assumptions section
        yPosition += 10;
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        doc.setFont('helvetica', 'bold');
        doc.text('Cost Estimation Assumptions:', 25, yPosition);
        yPosition += 6;
        
        const assumptions = [
            'Storage Cost: $0.02 per GB per month',
            'Compression Rate: 40% (data compression)',
            'Growth Rate: 20% (annual projection)',
            'Worker Node Cost: $8 per TB per month',
            'Retention Period: 1 year',
            'Daily Change Rate: 0.2%'
        ];
        
        assumptions.forEach(assumption => {
            if (yPosition < 280) {
                doc.setFontSize(9);
                doc.setTextColor(120, 120, 120);
                doc.setFont('helvetica', 'normal');
                doc.text(`• ${assumption}`, 30, yPosition);
                yPosition += 4;
            }
        });
    }
    
    return yPosition + 10;
}

function viewDetailedReport() {
    // Get the original HTML content that was uploaded
    if (window.dashboard && window.dashboard.originalHtmlContent) {
        const newWindow = window.open('', '_blank');
        newWindow.document.write(window.dashboard.originalHtmlContent);
        newWindow.document.close();
    } else {
        alert('Please upload an HTML report first to view the detailed analysis.');
    }
}

// Initialize dashboard when page loads
document.addEventListener('DOMContentLoaded', () => {
    new M365Dashboard();
});
