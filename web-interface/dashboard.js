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
                this.processFile(e.target.files[0]);
            }
        });

        // Click to upload
        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });
    }

    async processFile(file) {
        if (!file.name.endsWith('.html')) {
            alert('Please select an HTML report file.');
            return;
        }

        this.showLoading();
        
        try {
            const text = await file.text();
            this.parseReportData(text);
            this.displayDashboard();
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing the report file. Please ensure it\'s a valid HYCU M365 report.');
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
            backupRecommendations: this.extractBackupRecommendations(doc),
            costAnalysis: this.extractCostAnalysis(doc),
            teamsData: this.extractTeamsData(doc)
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

    extractBackupRecommendations(doc) {
        const recommendations = [];
        const recElements = doc.querySelectorAll('.recommendation');
        
        recElements.forEach(element => {
            const text = element.textContent.trim();
            if (text) {
                recommendations.push(text);
            }
        });

        return recommendations;
    }

    extractCostAnalysis(doc) {
        const monthlyStorage = this.extractNumberByLabel(doc, 'Monthly Storage Cost') || 0;
        const monthlyUser = this.extractNumberByLabel(doc, 'Monthly User Cost') || 0;
        const totalMonthly = this.extractNumberByLabel(doc, 'Total Monthly Cost') || 0;
        const annual = this.extractNumberByLabel(doc, 'Annual Cost') || 0;

        return { monthlyStorage, monthlyUser, totalMonthly, annual };
    }

    extractTeamsData(doc) {
        // This would need to be extracted from the actual report
        // For now, return default values
        return { totalTeams: 0, totalChannels: 0, avgChannels: 0 };
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
        this.populateBackupRecommendations();
        this.populateCostAnalysis();
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

    populateBackupRecommendations() {
        const container = document.getElementById('backup-recommendations');
        container.innerHTML = '';
        
        this.reportData.backupRecommendations.forEach(rec => {
            const div = document.createElement('div');
            div.className = 'recommendation';
            div.innerHTML = `<h4><i class="fas fa-check-circle"></i> ${rec}</h4>`;
            container.appendChild(div);
        });
    }

    populateCostAnalysis() {
        const data = this.reportData.costAnalysis;
        document.getElementById('monthly-storage-cost').textContent = `$${data.monthlyStorage.toFixed(2)}`;
        document.getElementById('monthly-user-cost').textContent = `$${data.monthlyUser.toFixed(2)}`;
        document.getElementById('total-monthly-cost').textContent = `$${data.totalMonthly.toFixed(2)}`;
        document.getElementById('annual-cost').textContent = `$${data.annual.toFixed(2)}`;
    }

    createCharts() {
        this.createStorageChart();
        this.createGrowthChart();
    }

    createStorageChart() {
        const ctx = document.getElementById('storage-chart').getContext('2d');
        const data = this.reportData.storageData;
        
        this.storageChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Exchange Online', 'OneDrive for Business', 'SharePoint Online'],
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
                        position: 'bottom'
                    }
                }
            }
        });
    }

    createGrowthChart() {
        const ctx = document.getElementById('growth-chart').getContext('2d');
        const data = this.reportData.growthData;
        
        const labels = Object.keys(data.projections).map(rate => `${rate}% Growth`);
        const values = Object.values(data.projections);
        
        this.growthChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Projected Size (GB)',
                    data: values,
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
                        beginAtZero: true
                    }
                }
            }
        });
    }
}

// Export functions
function exportToPDF() {
    alert('PDF export functionality would be implemented here using a library like jsPDF or Puppeteer.');
}

function exportToExcel() {
    alert('Excel export functionality would be implemented here using a library like SheetJS.');
}

function generateBackupPlan() {
    alert('Backup plan generation would create a detailed implementation plan based on the analysis.');
}

// Initialize dashboard when page loads
document.addEventListener('DOMContentLoaded', () => {
    new M365Dashboard();
});
