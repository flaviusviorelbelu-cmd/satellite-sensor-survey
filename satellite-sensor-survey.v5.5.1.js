<!DOCTYPE html>
<html lang="en">
<!-- v5.5.1: CRITICAL FIX 

-->
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Satellite Sensor Survey v5.5.1 - Form Detection Fix</title>
    <style>
        :root {
            --color-primary: #2180A3;
            --color-primary-hover: #1D6E80;
            --color-success: #32B8C6;
            --color-error: #C0152F;
            --color-warning: #A84B2F;
            --color-bg: #FCFCF9;
            --color-surface: #FFFFFF;
            --color-text: #134252;
            --color-text-secondary: #628076;
            --color-border: rgba(94, 82, 64, 0.2);
            --color-shadow: rgba(0, 0, 0, 0.08);
            --radius: 8px;
            --spacing: 16px;
        }

        * { box-sizing: border-box; }
        body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: var(--color-bg); color: var(--color-text); line-height: 1.5; }
        .container { max-width: 1400px; margin: 0 auto; padding: var(--spacing); }
        header { background: var(--color-surface); border-bottom: 1px solid var(--color-border); padding: var(--spacing) 0; margin-bottom: var(--spacing); }
        header h1 { margin: 0; font-size: 28px; color: var(--color-primary); }
        .header-subtitle { color: var(--color-text-secondary); font-size: 14px; margin-top: 4px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: var(--spacing); margin-bottom: var(--spacing); }
        .stat-card { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius); padding: var(--spacing); text-align: center; box-shadow: 0 2px 4px var(--color-shadow); }
        .stat-value { font-size: 32px; font-weight: 600; color: var(--color-primary); margin: 8px 0; }
        .stat-label { font-size: 12px; color: var(--color-text-secondary); text-transform: uppercase; letter-spacing: 0.5px; }
        .toolbar { display: flex; gap: var(--spacing); margin-bottom: var(--spacing); flex-wrap: wrap; align-items: center; }
        .search-box { flex: 1; min-width: 250px; display: flex; gap: 8px; }
        .search-box input { flex: 1; padding: 10px 12px; border: 1px solid var(--color-border); border-radius: var(--radius); font-size: 14px; background: var(--color-surface); color: var(--color-text); }
        .search-box input::placeholder { color: var(--color-text-secondary); }
        .search-box input:focus { outline: none; border-color: var(--color-primary); box-shadow: 0 0 0 3px rgba(33, 128, 141, 0.1); }
        .btn { padding: 10px 16px; border: none; border-radius: var(--radius); cursor: pointer; font-size: 14px; font-weight: 500; transition: all 150ms ease; display: inline-flex; align-items: center; gap: 8px; }
        .btn-primary { background: var(--color-primary); color: white; }
        .btn-primary:hover { background: var(--color-primary-hover); box-shadow: 0 4px 8px rgba(33, 128, 141, 0.2); }
        .btn-secondary { background: transparent; color: var(--color-primary); border: 1px solid var(--color-primary); }
        .btn-secondary:hover { background: rgba(33, 128, 141, 0.05); }
        .btn-small { padding: 6px 12px; font-size: 12px; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .content-grid { display: grid; grid-template-columns: 1fr 1fr; gap: var(--spacing); margin-bottom: var(--spacing); }
        @media (max-width: 1200px) { .content-grid { grid-template-columns: 1fr; } }
        .panel { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius); padding: var(--spacing); box-shadow: 0 2px 4px var(--color-shadow); }
        .panel-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: var(--spacing); border-bottom: 2px solid var(--color-border); padding-bottom: 12px; }
        .panel-header h2 { margin: 0; font-size: 18px; color: var(--color-primary); }
        .filter-group { display: flex; gap: 8px; margin-bottom: var(--spacing); flex-wrap: wrap; }
        .filter-btn { padding: 8px 12px; border: 1px solid var(--color-border); background: var(--color-surface); border-radius: var(--radius); cursor: pointer; font-size: 12px; transition: all 150ms ease; color: var(--color-text); }
        .filter-btn.active { background: var(--color-primary); color: white; border-color: var(--color-primary); }
        .filter-btn:hover { border-color: var(--color-primary); }
        .table-wrapper { overflow-x: auto; border-radius: var(--radius); border: 1px solid var(--color-border); }
        table { width: 100%; border-collapse: collapse; background: var(--color-surface); }
        thead { background: rgba(33, 128, 141, 0.08); border-bottom: 2px solid var(--color-border); }
        th { padding: 12px; text-align: left; font-weight: 600; font-size: 12px; color: var(--color-primary); text-transform: uppercase; letter-spacing: 0.5px; cursor: pointer; user-select: none; }
        th:hover { background: rgba(33, 128, 141, 0.12); }
        td { padding: 12px; border-bottom: 1px solid var(--color-border); font-size: 13px; }
        tbody tr { transition: background 150ms ease; }
        tbody tr:hover { background: rgba(33, 128, 141, 0.04); cursor: pointer; }
        tbody tr.selected { background: rgba(33, 128, 141, 0.1); }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .badge-operational { background: rgba(50, 184, 198, 0.2); color: var(--color-success); }
        .badge-testing { background: rgba(230, 129, 97, 0.2); color: var(--color-warning); }
        .badge-decommissioned { background: rgba(192, 21, 47, 0.2); color: var(--color-error); }
        .detail-view { display: grid; gap: 12px; }
        .detail-row { display: grid; grid-template-columns: 150px 1fr; gap: 12px; }
        .detail-label { font-weight: 600; color: var(--color-primary); font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
        .detail-value { color: var(--color-text); font-size: 14px; word-break: break-word; }
        .sensors-list { display: flex; flex-wrap: wrap; gap: 6px; padding-top: 6px; }
        .sensor-tag { background: rgba(33, 128, 141, 0.15); color: var(--color-primary); padding: 4px 10px; border-radius: 4px; font-size: 12px; display: inline-block; }
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.5); overflow: auto; }
        .modal.active { display: flex; align-items: center; justify-content: center; }
        .modal-content { background-color: var(--color-surface); padding: var(--spacing); border-radius: var(--radius); box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2); max-width: 600px; width: 90%; max-height: 90vh; overflow-y: auto; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: var(--spacing); padding-bottom: var(--spacing); border-bottom: 1px solid var(--color-border); }
        .modal-header h3 { margin: 0; color: var(--color-primary); }
        .modal h2 { margin-top: 0; margin-bottom: var(--spacing); color: var(--color-primary); }
        .modal-close { background: none; border: none; font-size: 24px; cursor: pointer; color: var(--color-text-secondary); }
        .modal-close:hover { color: var(--color-text); }
        .form-group { margin-bottom: var(--spacing); }
        .form-group label { display: block; margin-bottom: 6px; font-weight: 600; color: var(--color-text); font-size: 13px; }
        .form-group.required label::after { content: ' *'; color: var(--color-error); }
        .form-group input, .form-group textarea, .form-group select { width: 100%; padding: 10px 12px; border: 1px solid var(--color-border); border-radius: var(--radius); font-size: 13px; font-family: inherit; color: var(--color-text); background: var(--color-surface); }
        .form-group input:focus, .form-group textarea:focus, .form-group select:focus { outline: none; border-color: var(--color-primary); box-shadow: 0 0 0 3px rgba(33, 128, 141, 0.1); }
        .form-group input.error, .form-group textarea.error, .form-group select.error { border-color: var(--color-error); background-color: rgba(192, 21, 47, 0.05); }
        .form-group textarea { resize: vertical; min-height: 100px; }
        .form-actions { display: flex; gap: 12px; margin-top: var(--spacing); }
        .form-actions button { flex: 1; }
        .modal-footer { display: flex; gap: 8px; justify-content: flex-end; margin-top: var(--spacing); padding-top: var(--spacing); border-top: 1px solid var(--color-border); }
        .empty-state { text-align: center; padding: 40px var(--spacing); color: var(--color-text-secondary); }
        .empty-state-icon { font-size: 48px; margin-bottom: 12px; }
        .loading { display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(33, 128, 141, 0.2); border-top-color: var(--color-primary); border-radius: 50%; animation: spin 0.8s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .alert { padding: 12px 16px; border-radius: var(--radius); margin-bottom: var(--spacing); display: flex; align-items: center; gap: 12px; animation: slideIn 300ms ease-out; }
        @keyframes slideIn { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
        .alert-success { background: rgba(50, 184, 198, 0.1); color: var(--color-success); border: 1px solid rgba(50, 184, 198, 0.3); }
        .alert-error { background: rgba(192, 21, 47, 0.1); color: var(--color-error); border: 1px solid rgba(192, 21, 47, 0.3); }
        .alert-info { background: rgba(33, 128, 141, 0.1); color: var(--color-primary); border: 1px solid rgba(33, 128, 141, 0.3); }
        .alert-warning { background: rgba(230, 129, 97, 0.1); color: var(--color-warning); border: 1px solid rgba(230, 129, 97, 0.3); }
        .tabs { display: flex; border-bottom: 2px solid var(--color-border); margin-bottom: var(--spacing); gap: 4px; }
        .tab-btn { padding: 12px 16px; background: none; border: none; cursor: pointer; color: var(--color-text-secondary); font-weight: 500; font-size: 14px; border-bottom: 3px solid transparent; transition: all 150ms ease; }
        .tab-btn.active { color: var(--color-primary); border-bottom-color: var(--color-primary); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .validation-error { color: var(--color-error); font-size: 12px; margin-top: 4px; display: block; min-height: 16px; }
        .storage-indicator { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; margin-left: 8px; background: rgba(50, 184, 198, 0.15); color: var(--color-success); }
        .csv-info { background: rgba(33, 128, 141, 0.05); border: 1px solid rgba(33, 128, 141, 0.2); border-radius: var(--radius); padding: 12px; margin-bottom: var(--spacing); font-size: 12px; color: var(--color-text-secondary); }
        .csv-info strong { color: var(--color-primary); }
        .import-preview { background: var(--color-bg); border: 1px solid var(--color-border); border-radius: var(--radius); padding: 12px; margin-top: 12px; max-height: 300px; overflow-y: auto; font-size: 12px; }
        .import-preview table { font-size: 12px; }
        .import-preview th, .import-preview td { padding: 6px; }
        .action-buttons { display: flex; gap: 8px; }
        .action-btn { padding: 4px 8px; font-size: 11px; white-space: nowrap; }
    </style>
</head>
<body>
    <header>
        <div class="container">
            <h1>🛰️ Satellite Sensor Survey v5.4.9</h1>
            <p class="header-subtitle">Form detection fix - Critical bug resolved <span class="storage-indicator" id="storageIndicator">Local Storage Mode</span></p>
        </div>
    </header>

    <main class="container">
        <div id="alertContainer"></div>
        <div class="stats-grid">
            <div class="stat-card"><div class="stat-label">Total Satellites</div><div class="stat-value" id="statSatellites">0</div></div>
            <div class="stat-card"><div class="stat-label">Total Sensors</div><div class="stat-value" id="statSensors">0</div></div>
            <div class="stat-card"><div class="stat-label">Operational</div><div class="stat-value" id="statOperational">0</div></div>
            <div class="stat-card"><div class="stat-label">Constellations</div><div class="stat-value" id="statConstellations">0</div></div>
        </div>
        <div class="toolbar">
            <div class="search-box"><input type="text" id="searchInput" placeholder="Search..."></div>
            <button type="button" class="btn btn-primary" id="addSatelliteBtn">+ Add Satellite</button>
            <button type="button" class="btn btn-secondary" id="importBtn">📤 Import CSV</button>
            <button type="button" class="btn btn-secondary" id="exportBtn">📥 Export</button>
        </div>
        <div class="content-grid">
            <div class="panel"><div class="panel-header"><h2>Satellites</h2><span id="satelliteCount">0</span></div><div class="tabs"><button type="button" class="tab-btn active" onclick="switchTab('satellites', 'sat-all'); return false;">All</button><button type="button" class="tab-btn" onclick="switchTab('satellites', 'sat-operational'); return false;">Operational</button><button type="button" class="tab-btn" onclick="switchTab('satellites', 'sat-leo'); return false;">LEO</button><button type="button" class="tab-btn" onclick="switchTab('satellites', 'sat-geo'); return false;">GEO</button></div><div class="table-wrapper"><table id="satelliteTable"><thead><tr><th onclick="sortTable('satellites', 'Title')">Name ↕</th><th onclick="sortTable('satellites', 'NORAD_ID')">NORAD ID ↕</th><th>Status</th><th>Orbit</th><th style="width: 80px; text-align: center;">Actions</th></tr></thead><tbody id="satelliteTableBody"><tr><td colspan="5" style="text-align: center;"><div class="loading"></div></td></tr></tbody></table></div></div>
            <div class="panel"><div class="panel-header"><h2>Sensors</h2><span id="sensorCount">0</span></div><div class="tabs"><button type="button" class="tab-btn active" onclick="switchTab('sensors', 'sen-all'); return false;">All</button><button type="button" class="tab-btn" onclick="switchTab('sensors', 'sen-eo'); return false;">EO</button><button type="button" class="tab-btn" onclick="switchTab('sensors', 'sen-ir'); return false;">IR</button><button type="button" class="tab-btn" onclick="switchTab('sensors', 'sen-sar'); return false;">SAR</button></div><div class="table-wrapper"><table id="sensorTable"><thead><tr><th onclick="sortTable('sensors', 'Title')">Name ↕</th><th>Type</th><th>Satellites</th></tr></thead><tbody id="sensorTableBody"><tr><td colspan="3" style="text-align: center;"><div class="loading"></div></td></tr></tbody></table></div></div>
        </div>
        <div class="panel"><div class="panel-header"><h2 id="detailTitle">Satellite Details</h2><div id="detailActions"></div></div><div id="detailContent" class="empty-state"><div class="empty-state-icon">📋</div><p>Select a satellite or sensor</p></div></div>
    </main>

    <div id="addSatelliteModal" class="modal" onclick="handleModalBackdropClick(event); return false;">
        <div class="modal-content" onclick="event.stopPropagation();">
            <div class="modal-header"><h3 id="modalTitle">Add New Satellite</h3><button type="button" class="modal-close" onclick="closeAddSatelliteModal(); return false;">×</button></div>
            <form id="satelliteForm">
                <div class="form-group required"><label for="satellite-title">Satellite Name</label><input type="text" id="satellite-title" name="Title" placeholder="Enter satellite name" maxlength="255"><div class="validation-error" id="error-title"></div></div>
                <div class="form-group required"><label for="satellite-norad">NORAD ID</label><input type="text" id="satellite-norad" name="NORAD_ID" placeholder="Enter NORAD ID" pattern="[0-9]*" maxlength="10"><div class="validation-error" id="error-norad"></div></div>
                <div class="form-group"><label for="satellite-cospar">COSPAR ID</label><input type="text" id="satellite-cospar" name="COSPAR_ID" placeholder="e.g., 2013-008A" maxlength="20"><div class="validation-error" id="error-cospar"></div></div>
                <div class="form-group"><label for="satellite-mission">Mission Type</label><input type="text" id="satellite-mission" name="Mission_Type" placeholder="e.g., Earth Observation" maxlength="100"></div>
                <div class="form-group"><label for="satellite-status">Status</label><select id="satellite-status" name="Status"><option value="Operational">Operational</option><option value="Non-operational">Non-operational</option><option value="Partially operational">Partially operational</option><option value="Unknown">Unknown</option></select></div>
                <div class="form-group"><label for="satellite-orbit">Orbit Type</label><select id="satellite-orbit" name="Orbit_Type"><option value="">-- Select --</option><option value="LEO">LEO</option><option value="MEO">MEO</option><option value="GEO">GEO</option><option value="SSO">SSO</option></select></div>
                <div class="form-group"><label for="satellite-launch">Launch Date</label><input type="date" id="satellite-launch" name="Launch_Date"><div class="validation-error" id="error-launch"></div></div>
                <div class="form-group"><label for="satellite-sensors">Sensor Names</label><textarea id="satellite-sensors" name="Sensor_Names" rows="3" placeholder="Comma-separated" maxlength="500"></textarea></div>
                <div class="form-actions"><button type="button" class="btn btn-primary" id="saveSatelliteBtn">Save Satellite</button><button type="button" class="btn btn-secondary" onclick="closeAddSatelliteModal(); return false;">Cancel</button></div>
            </form>
        </div>
    </div>

    <div id="importModal" class="modal" onclick="handleImportModalBackdropClick(event); return false;">
        <div class="modal-content" onclick="event.stopPropagation();">
            <div class="modal-header"><h3>Import CSV</h3><button type="button" class="modal-close" onclick="closeImportModal(); return false;">×</button></div>
            <div class="csv-info"><strong>CSV Format:</strong> Title, NORAD_ID, COSPAR_ID, Mission_Type, Status, Orbit_Type, Launch_Date, Sensor_Names</div>
            <form id="importForm"><div class="form-group"><label for="csvFile">Choose CSV File</label><input type="file" id="csvFile" name="csvFile" accept=".csv"><div class="validation-error" id="error-csv"></div></div><div id="importPreview" class="import-preview" style="display:none;"></div><div class="form-actions"><button type="submit" class="btn btn-primary">Import Data</button><button type="button" class="btn btn-secondary" onclick="closeImportModal(); return false;">Cancel</button></div></form>
        </div>
    </div>

    <script>
/**
 * SATELLITE SENSOR SURVEY v5.5.1
 * 
 * Enhanced version with integrated SatelliteFormHandler for proper form management
 * Fixes: FormData construction error, form validation, SharePoint REST API integration
 * 
 * @version 5.5.1
 * @author DevOps Team - NCIA
 * @date 2026-01-20
 * 
 * CHANGELOG v5.5.1:
 * - Fixed TypeError: Failed to construct 'FormData': parameter 1 is not of type 'HTMLFormElement'
 * - Changed saveSatellite to use validation.data directly instead of FormData constructor
 * - Updated saveSatelliteToSharePoint to access form fields as object properties
 * - Updated saveSatelliteToLocalStorage to access form fields as object properties
 * - Improved null-safety for form data access throughout save pipeline
 */

// ============================================================================
// POLYFILL: AbortSignal.timeout
// ============================================================================

if (!AbortSignal.timeout) {
    AbortSignal.timeout = function(ms) {
        const controller = new AbortController();
        setTimeout(() => controller.abort(), ms);
        return controller.signal;
    };
}

// ============================================================================
// SATELLITE FORM HANDLER CLASS - v1.0.0
// ============================================================================

class SatelliteFormHandler {
  constructor(options = {}) {
    this.config = {
      siteUrl: options.siteUrl || (typeof _spPageContextInfo !== 'undefined' ? _spPageContextInfo.webAbsoluteUrl : ''),
      listTitle: options.listTitle || 'Satellite_Fixed',
      logLevel: options.logLevel || 'INFO',
      ...options
    };

    this.fieldMap = {
      Title: 'satellite-title',
      NORAD_ID: 'satellite-norad',
      COSPAR_ID: 'satellite-cospar',
      Mission_Type: 'satellite-mission',
      Status: 'satellite-status',
      Orbit_Type: 'satellite-orbit',
      Launch_Date: 'satellite-launch',
      Sensor_Names: 'satellite-sensors'
    };

    this.currentEditId = null;
    this.logger = this._initLogger();
    this._validateSetup();
  }

  _initLogger() {
    const levels = { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3 };
    const currentLevel = levels[this.config.logLevel] || 1;

    return {
      debug: (msg, data) => currentLevel <= 0 && console.log(`[DEBUG] 🔧 ${msg}`, data || ''),
      info: (msg, data) => currentLevel <= 1 && console.log(`[INFO] ℹ️ ${msg}`, data || ''),
      warn: (msg, data) => currentLevel <= 2 && console.warn(`[WARN] ⚠️ ${msg}`, data || ''),
      error: (msg, data) => currentLevel <= 3 && console.error(`[ERROR] ❌ ${msg}`, data || '')
    };
  }

  _validateSetup() {
    const missingFields = [];
    Object.entries(this.fieldMap).forEach(([sp, htmlId]) => {
      if (!document.getElementById(htmlId)) {
        missingFields.push(`${sp} (${htmlId})`);
      }
    });

    if (missingFields.length > 0) {
      this.logger.warn(`Missing form fields: ${missingFields.join(', ')}`);
    } else {
      this.logger.info('✅ All form fields detected');
    }

    if (typeof _spPageContextInfo === 'undefined') {
      this.logger.warn('SharePoint context not available');
    }
  }

  getForm() {
    let form = document.getElementById('satelliteForm');
    if (form) {
      this.logger.debug('Form found via getElementById');
      return form;
    }

    const modal = document.getElementById('addSatelliteModal');
    if (modal) {
      form = modal.querySelector('#satelliteForm');
      if (form) {
        this.logger.debug('Form found via modal querySelector');
        return form;
      }
    }

    form = document.querySelector('form#satelliteForm');
    if (form) {
      this.logger.debug('Form found via document.querySelector');
      return form;
    }

    this.logger.error('Form not found using any strategy');
    return null;
  }

  getFormInput(inputId) {
    let input = document.getElementById(inputId);
    if (input) return input;

    const modal = document.getElementById('addSatelliteModal');
    if (modal) {
      input = modal.querySelector('#' + inputId);
      if (input) return input;
    }

    return document.querySelector('#' + inputId);
  }

  getFormData() {
    const data = {};
    Object.entries(this.fieldMap).forEach(([spField, htmlId]) => {
      const element = this.getFormInput(htmlId);
      if (element) {
        data[spField] = element.value || '';
      }
    });
    this.logger.debug('Form data retrieved', data);
    return data;
  }

  setFormData(data) {
    Object.entries(data).forEach(([spField, value]) => {
      const htmlId = this.fieldMap[spField];
      const element = this.getFormInput(htmlId);
      if (element) {
        element.value = value || '';
        this.logger.debug(`Set ${htmlId} = ${value}`);
      }
    });
  }

  clearForm() {
    Object.values(this.fieldMap).forEach(htmlId => {
      const element = this.getFormInput(htmlId);
      if (element) {
        element.value = '';
      }
    });
    this.currentEditId = null;
    this.logger.info('Form cleared');
  }

  validateForm() {
    const data = this.getFormData();
    const errors = [];

    if (!data.Title || data.Title.trim() === '') {
      errors.push('Satellite Title is required');
    }
    if (!data.NORAD_ID || data.NORAD_ID.trim() === '') {
      errors.push('NORAD ID is required');
    }
    if (!data.Status || data.Status.trim() === '') {
      errors.push('Status is required');
    }

    if (data.NORAD_ID && !/^\d+$/.test(data.NORAD_ID)) {
      errors.push('NORAD ID must be numeric');
    }

    if (errors.length > 0) {
      this.logger.warn('Validation failed', errors);
      return { valid: false, errors };
    }

    this.logger.debug('Validation passed', data);
    return { valid: true, errors: [], data };
  }

  async saveSatellite() {
    this.logger.info('💾 saveSatellite called');

    const validation = this.validateForm();
    if (!validation.valid) {
      this.logger.error('Form validation failed', validation.errors);
      return { success: false, errors: validation.errors };
    }

    try {
      const data = validation.data;
      const isNewItem = !this.currentEditId;
      
      return {
        success: true,
        data: data,
        isNewItem: isNewItem,
        editingId: this.currentEditId
      };

    } catch (error) {
      this.logger.error('Save preparation failed', error.message);
      return { success: false, errors: [error.message] };
    }
  }

  async editSatellite(satelliteId) {
    this.logger.info(`✏️ editSatellite called with ID: ${satelliteId}`);
    this.currentEditId = satelliteId;
    return true;
  }

  addSatellite() {
    this.logger.info('📝 Opening add mode');
    this.currentEditId = null;
    this.clearForm();
  }

  debug() {
    console.group('🔍 SatelliteFormHandler Debug Info');
    console.log('Configuration:', this.config);
    console.log('Current Edit ID:', this.currentEditId);
    console.log('Form Data:', this.getFormData());
    console.log('Field Map:', this.fieldMap);
    Object.values(this.fieldMap).forEach(htmlId => {
      const el = this.getFormInput(htmlId);
      console.log(`  ${htmlId}:`, el ? { exists: true, value: el.value } : { exists: false });
    });
    console.groupEnd();
  }
}

// ============================================================================
// CONFIGURATION & STATE
// ============================================================================

const config = {
    satelliteList: 'Satellite_Fixed',
    sensorList: 'Sensor',
    inSharePoint: typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo?.webAbsoluteUrl ? true : false,
    siteUrl: typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo?.webAbsoluteUrl ? _spPageContextInfo.webAbsoluteUrl : '',
    apiTimeout: 30000
};

const state = {
    satellites: [],
    sensors: [],
    filteredSatellites: [],
    filteredSensors: [],
    selectedSatellite: null,
    selectedSensor: null,
    searchTerm: '',
    sortColumn: 'Title',
    sortDirection: 'asc',
    currentFilter: 'all',
    editingId: null,
    csvPreviewData: [],
    nextId: 1000,
    isLoading: false,
    isPendingSave: false
};

// ============================================================================
// LOGGER SYSTEM
// ============================================================================

const Logger = {
    levels: { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3 },
    currentLevel: 1,
    log: function(level, message, data = null) {
        if (level < this.currentLevel) return;
        const timestamp = new Date().toLocaleTimeString();
        const levelName = Object.keys(this.levels).find(key => this.levels[key] === level);
        const prefix = `[${timestamp}] [${levelName}]`;
        if (data) console.log(prefix, message, data); else console.log(prefix, message);
    },
    debug: function(msg, data) { this.log(this.levels.DEBUG, msg, data); },
    info: function(msg, data) { this.log(this.levels.INFO, msg, data); },
    warn: function(msg, data) { this.log(this.levels.WARN, msg, data); },
    error: function(msg, data) { this.log(this.levels.ERROR, msg, data); }
};

Logger.info('🚀 Initializing Satellite Sensor Survey v5.5.1');
Logger.info('SharePoint Mode:', config.inSharePoint ? 'YES' : 'NO');

// Initialize form handler
const formHandler = new SatelliteFormHandler({
    siteUrl: config.siteUrl,
    listTitle: config.satelliteList,
    logLevel: 'INFO'
});

// ============================================================================
// LOCAL STORAGE FUNCTIONS
// ============================================================================

function saveToLocalStorage() {
    try {
        localStorage.setItem('satellites_data', JSON.stringify(state.satellites));
        localStorage.setItem('nextId', state.nextId.toString());
        Logger.info('💾 Data saved');
    } catch (error) {
        if (error.name === 'QuotaExceededError') {
            Logger.error('Storage quota exceeded');
            showAlert('⚠️ Storage full', 'warning');
        } else {
            Logger.error('Save error:', { error: error.message });
            showAlert('⚠️ Could not save', 'warning');
        }
    }
}

function loadFromLocalStorage() {
    try {
        const stored = localStorage?.getItem('satellites_data');
        if (stored) {
            state.satellites = JSON.parse(stored) ?? [];
            Logger.info('📁 Data loaded', { count: state.satellites.length });
        }
        const storedId = localStorage?.getItem('nextId');
        state.nextId = storedId ? parseInt(storedId) : Math.max(1000, ...(state.satellites?.map(s => s?.ID ?? 0) ?? [])) + 1;
    } catch (error) {
        Logger.error('Load error:', { error: error.message });
        state.satellites = [];
    }
}

// ============================================================================
// EVENT LISTENERS
// ============================================================================

function attachEventListeners() {
    Logger.info('🔗 Attaching event listeners...');
    try {
        document.addEventListener('submit', function(event) {
            if (event?.target?.id === 'satelliteForm') {
                event.preventDefault();
                event.stopPropagation();
                event.stopImmediatePropagation();
                Logger.info('📋 Form submit captured');
                saveSatellite(event);
                return false;
            }
        }, true);

        document.addEventListener('click', function(event) {
            if (event?.target?.id === 'saveSatelliteBtn') {
                event.preventDefault();
                event.stopPropagation();
                event.stopImmediatePropagation();
                Logger.info('📋 Save button clicked');
                saveSatellite(event);
                return false;
            }
        }, true);

        document.addEventListener('click', (e) => {
            const editBtn = e.target?.closest?.('.edit-btn');
            const deleteBtn = e.target?.closest?.('.delete-btn');
            
            if (editBtn) {
                e.preventDefault();
                const id = parseInt(editBtn.dataset?.id);
                if (!isNaN(id)) {
                    Logger.info('✏️ Edit clicked for ID: ' + id);
                    editSatellite(id);
                }
                return false;
            }
            
            if (deleteBtn) {
                e.preventDefault();
                const id = parseInt(deleteBtn.dataset?.id);
                if (!isNaN(id)) {
                    Logger.info('🗑️ Delete clicked for ID: ' + id);
                    deleteSatellite(id);
                }
                return false;
            }
        });

        document.getElementById('addSatelliteBtn')?.addEventListener('click', (e) => { e.preventDefault(); showAddSatelliteModal(); });
        document.getElementById('importBtn')?.addEventListener('click', (e) => { e.preventDefault(); showImportModal(); });
        document.getElementById('exportBtn')?.addEventListener('click', (e) => { e.preventDefault(); exportToCSV(); });
        document.getElementById('searchInput')?.addEventListener('input', (e) => { state.searchTerm = (e.target?.value ?? '').toLowerCase(); filterAndDisplayData(); });
        document.getElementById('csvFile')?.addEventListener('change', handleCSVFileSelect);
        document.getElementById('importForm')?.addEventListener('submit', handleCSVImport);

        Logger.info('✅ Event listeners attached');
    } catch (error) {
        Logger.error('Failed to attach listeners:', { error: error.message });
    }
}

// ============================================================================
// INITIALIZATION
// ============================================================================

document.addEventListener('DOMContentLoaded', async () => {
    Logger.info('📋 Page loaded');
    try {
        const indicator = document.getElementById('storageIndicator');
        if (indicator) {
            indicator.textContent = config.inSharePoint ? 'SharePoint Mode (Connected)' : 'Local Storage Mode';
        }

        if (config.inSharePoint && config.siteUrl) {
            showAlert('🔄 Loading from SharePoint...', 'info');
            await loadSatellites();
        } else {
            loadFromLocalStorage();
            if ((state.satellites?.length ?? 0) === 0) {
                state.satellites = [{ ID: 1, Title: 'Landsat 8', NORAD_ID: '39084', COSPAR_ID: '2013-008A', Mission_Type: 'Earth Observation', Status: 'Operational', Orbit_Type: 'SSO', Launch_Date: '2013-02-11', Constellation_ID: 1, Sensor_Names: 'Multispectral Imager' }];
                state.nextId = 1001;
                saveToLocalStorage();
            }
        }

        await loadSensors();
        attachEventListeners();
        filterAndDisplayData();
        updateStatistics();
        Logger.info('✅ Init complete');
    } catch (error) {
        Logger.error('Init failed:', { error: error.message });
        showAlert('❌ Initialization failed', 'error');
    }
});

// ============================================================================
// API FUNCTIONS
// ============================================================================

async function loadSatellites() {
    if (!config.inSharePoint || !config.siteUrl) return;
    state.isLoading = true;
    try {
        const url = config.siteUrl + `/_api/web/lists/getbytitle('${config.satelliteList}')/items?$select=ID,Title,NORAD_ID,COSPAR_ID,Mission_Type,Status,Orbit_Type,Launch_Date,Expected_Lifetime,Constellation_ID,Sensor_Names,Primary_Sensor&$top=5000`;
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json;odata=verbose' },
            credentials: 'include',
            signal: AbortSignal.timeout(config.apiTimeout)
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();
        state.satellites = (data?.d?.results ?? []).map((item, idx) => ({
            ID: item?.ID,
            Title: item?.Title ?? `Satellite ${idx}`,
            NORAD_ID: item?.NORAD_ID?.toString?.() ?? '',
            COSPAR_ID: item?.COSPAR_ID ?? '',
            Mission_Type: item?.Mission_Type ?? '',
            Status: item?.Status ?? 'Operational',
            Orbit_Type: item?.Orbit_Type ?? '',
            Launch_Date: item?.Launch_Date ? new Date(item.Launch_Date).toISOString().split('T')[0] : '',
            Expected_Lifetime: item?.Expected_Lifetime ?? '',
            Constellation_ID: item?.Constellation_ID ?? '',
            Sensor_Names: item?.Sensor_Names ?? '',
            Primary_Sensor: item?.Primary_Sensor ?? ''
        })).filter(item => item !== null);
        Logger.info('✅ Satellites loaded', { count: state.satellites.length });
        showAlert(`✅ Loaded ${state.satellites.length} satellites`, 'success');
    } catch (error) {
        Logger.error('Error loading satellites:', { error: error.message });
        showAlert(`❌ Error: ${error.message}`, 'error');
    } finally {
        state.isLoading = false;
    }
}

async function loadSensors() {
    if (!config.inSharePoint || !config.siteUrl) { state.sensors = []; return; }
    try {
        const url = config.siteUrl + `/_api/web/lists/getbytitle('${config.sensorList}')/items?$select=ID,Title,Sensor_Type,Description,N_of_Bands,Resolution_Min_m,Resolution_Max_m,Swath_Min_km,Swath_Max_km,Satellite_Names&$top=5000`;
        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json;odata=verbose' },
            credentials: 'include',
            signal: AbortSignal.timeout(config.apiTimeout)
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();
        state.sensors = (data?.d?.results ?? []).map(item => ({
            ID: item?.ID,
            Title: item?.Title ?? 'Unknown',
            Sensor_Type: item?.Sensor_Type ?? '',
            Description: item?.Description ?? '',
            Satellite_Names: item?.Satellite_Names ?? ''
        }));
        Logger.info('✅ Sensors loaded', { count: state.sensors.length });
    } catch (error) {
        Logger.error('Error loading sensors:', { error: error.message });
        state.sensors = [];
    }
}

// ============================================================================
// SAVE FUNCTIONS - v5.5.1 FIXES
// ============================================================================

function saveSatellite(event) {
    if (state.isPendingSave) { Logger.warn('Save already pending'); return false; }
    event?.preventDefault?.();
    Logger.info('💾 saveSatellite called');
    clearValidationErrors();
    state.isPendingSave = true;

    try {
        const validation = formHandler.validateForm();
        if (!validation.valid) {
            Logger.error('Form validation failed', validation.errors);
            validation.errors.forEach((error, idx) => {
                showValidationError('title', error);
            });
            showAlert('❌ Fix validation errors', 'error');
            return false;
        }

        Logger.info('✅ Form validated');
        const formData = validation.data;
        
        if (config.inSharePoint && config.siteUrl) {
            saveSatelliteToSharePoint(formData);
        } else {
            saveSatelliteToLocalStorage(formData);
        }
    } finally {
        state.isPendingSave = false;
    }
    return false;
}

async function saveSatelliteToSharePoint(formData) {
    try {
        const payload = {
            '__metadata': { 'type': 'SP.Data.Satellite_x005f_FixedListItem' },
            'Title': (formData.Title ?? '')?.trim?.() ?? '',
            'NORAD_ID': parseInt((formData.NORAD_ID ?? '')?.trim?.() ?? 0),
            'COSPAR_ID': (formData.COSPAR_ID ?? '')?.trim?.() ?? '',
            'Mission_Type': (formData.Mission_Type ?? '')?.trim?.() ?? '',
            'Status': formData.Status ?? 'Operational',
            'Orbit_Type': formData.Orbit_Type ?? '',
            'Launch_Date': formData.Launch_Date ? new Date(formData.Launch_Date).toISOString() : null,
            'Sensor_Names': (formData.Sensor_Names ?? '')?.trim?.() ?? ''
        };

        const contextUrl = config.siteUrl + '/_api/contextinfo';
        const digestResponse = await fetch(contextUrl, {
            method: 'POST',
            headers: { 'Accept': 'application/json;odata=verbose' },
            credentials: 'include',
            signal: AbortSignal.timeout(config.apiTimeout)
        });

        if (!digestResponse.ok) throw new Error('Failed to get digest');
        const digestData = await digestResponse.json();
        const digest = digestData?.d?.GetContextWebInformation?.FormDigestValue;

        let url, method, headers;
        if (state.editingId) {
            url = config.siteUrl + `/_api/web/lists/getbytitle('${config.satelliteList}')/items(${state.editingId})`;
            method = 'POST';
            headers = { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };
        } else {
            url = config.siteUrl + `/_api/web/lists/getbytitle('${config.satelliteList}')/items`;
            method = 'POST';
            headers = { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest };
        }

        const response = await fetch(url, { method, headers, body: JSON.stringify(payload), credentials: 'include', signal: AbortSignal.timeout(config.apiTimeout) });
        if (response.ok || response.status === 201 || response.status === 204) {
            const action = state.editingId ? 'updated' : 'added';
            Logger.info(`✅ Satellite ${action}`);
            showAlert(`✅ Satellite ${action}!`, 'success');
            closeAddSatelliteModal();
            state.editingId = null;
            await loadSatellites();
            filterAndDisplayData();
        } else {
            throw new Error(`HTTP ${response.status}`);
        }
    } catch (error) {
        Logger.error('SharePoint save error:', { error: error.message });
        showAlert('❌ Error: ' + error.message, 'error');
    }
}

function saveSatelliteToLocalStorage(formData) {
    try {
        const titleValue = (formData.Title ?? '')?.trim?.() ?? '';
        const noradValue = (formData.NORAD_ID ?? '')?.trim?.() ?? '';
        
        if (state.editingId) {
            const index = state.satellites?.findIndex(s => s?.ID === state.editingId);
            if (index >= 0) {
                state.satellites[index] = { 
                    ...state.satellites[index], 
                    Title: titleValue, 
                    NORAD_ID: noradValue, 
                    COSPAR_ID: (formData.COSPAR_ID ?? '')?.trim?.() ?? '', 
                    Mission_Type: (formData.Mission_Type ?? '')?.trim?.() ?? '', 
                    Status: formData.Status ?? 'Operational', 
                    Orbit_Type: formData.Orbit_Type ?? '', 
                    Launch_Date: formData.Launch_Date ?? '', 
                    Sensor_Names: (formData.Sensor_Names ?? '')?.trim?.() ?? '' 
                };
                Logger.info('✅ Satellite updated');
                showAlert('✅ Updated!', 'success');
            }
        } else {
            state.satellites.push({ 
                ID: state.nextId++, 
                Title: titleValue, 
                NORAD_ID: noradValue, 
                COSPAR_ID: (formData.COSPAR_ID ?? '')?.trim?.() ?? '', 
                Mission_Type: (formData.Mission_Type ?? '')?.trim?.() ?? '', 
                Status: formData.Status ?? 'Operational', 
                Orbit_Type: formData.Orbit_Type ?? '', 
                Launch_Date: formData.Launch_Date ?? '', 
                Expected_Lifetime: '', 
                Constellation_ID: 1, 
                Sensor_Names: (formData.Sensor_Names ?? '')?.trim?.() ?? '', 
                Primary_Sensor: '' 
            });
            Logger.info('✅ Satellite added');
            showAlert('✅ Added!', 'success');
        }
        saveToLocalStorage();
        closeAddSatelliteModal();
        state.editingId = null;
        filterAndDisplayData();
        updateStatistics();
    } catch (error) {
        Logger.error('LocalStorage save error:', { error: error.message });
        showAlert('❌ Error: ' + error.message, 'error');
    }
}

// ============================================================================
// UI FUNCTIONS
// ============================================================================

function filterAndDisplayData() {
    try {
        state.filteredSatellites = (state.satellites ?? []).filter(sat => {
            const matchesSearch = !state.searchTerm || (sat?.Title?.toLowerCase?.()?.includes(state.searchTerm)) || (sat?.NORAD_ID?.toString?.()?.includes(state.searchTerm));
            if (state.currentFilter === 'sat-operational') return matchesSearch && sat?.Status === 'Operational';
            if (state.currentFilter === 'sat-leo') return matchesSearch && sat?.Orbit_Type === 'LEO';
            if (state.currentFilter === 'sat-geo') return matchesSearch && sat?.Orbit_Type === 'GEO';
            return matchesSearch;
        });
        renderSatelliteTable();
        document.getElementById('satelliteCount').textContent = state.filteredSatellites.length;
    } catch (error) {
        Logger.error('Filter error:', { error: error.message });
    }
}

function renderSatelliteTable() {
    const tbody = document.getElementById('satelliteTableBody');
    if (!tbody) return;
    if (state.filteredSatellites.length === 0) { tbody.innerHTML = '<tr><td colspan="5" class="empty-state">No satellites</td></tr>'; return; }
    tbody.innerHTML = state.filteredSatellites.map(sat => {
        const statusClass = (sat?.Status ?? '').toLowerCase().replace(/[- ]/g, '');
        return `<tr onclick="selectSatellite(${sat?.ID})" class="${state.selectedSatellite?.ID === sat?.ID ? 'selected' : ''}"><td><strong>${escapeHtml(sat?.Title ?? '')}</strong></td><td>${escapeHtml(sat?.NORAD_ID ?? '')}</td><td><span class="badge badge-${statusClass}">${escapeHtml(sat?.Status ?? '')}</span></td><td>${escapeHtml(sat?.Orbit_Type ?? '')}</td><td style="text-align: center;"><div class="action-buttons" style="justify-content: center;"><button class="btn btn-small btn-secondary edit-btn action-btn" data-id="${sat?.ID}">✏️</button><button class="btn btn-small btn-secondary delete-btn action-btn" data-id="${sat?.ID}">🗑️</button></div></td></tr>`;
    }).join('');
}

function showAddSatelliteModal() {
    Logger.info('📝 Opening add modal');
    state.editingId = null;
    formHandler.addSatellite();
    const modal = document.getElementById('addSatelliteModal');
    if (!modal) return false;
    modal.classList.add('active');
    document.getElementById('modalTitle').textContent = 'Add New Satellite';
    setTimeout(() => {
        clearValidationErrors();
        document.getElementById('satellite-title')?.focus();
    }, 50);
    return false;
}

function closeAddSatelliteModal() {
    const modal = document.getElementById('addSatelliteModal');
    if (modal) modal.classList.remove('active');
    formHandler.clearForm();
    clearValidationErrors();
    state.editingId = null;
    return false;
}

function handleModalBackdropClick(event) {
    if (event.target.id === 'addSatelliteModal') closeAddSatelliteModal();
    return false;
}

function handleImportModalBackdropClick(event) {
    if (event.target.id === 'importModal') closeImportModal();
    return false;
}

function clearValidationErrors() {
    document.querySelectorAll('.validation-error').forEach(el => el.textContent = '');
    document.querySelectorAll('.form-group input, .form-group textarea, .form-group select').forEach(input => input.classList.remove('error'));
}

function showValidationError(fieldName, message) {
    const errorEl = document.getElementById(`error-${fieldName}`);
    const inputEl = formHandler.getFormInput(`satellite-${fieldName}`);
    if (errorEl) errorEl.textContent = message;
    if (inputEl) inputEl.classList.add('error');
}

function editSatellite(id) {
    Logger.info('📝 editSatellite with ID: ' + id);
    const sat = state.satellites?.find(s => s?.ID === id);
    if (!sat) { Logger.error('Satellite not found'); showAlert('❌ Not found', 'error'); return false; }
    
    state.editingId = id;
    formHandler.editSatellite(id);
    const modal = document.getElementById('addSatelliteModal');
    if (!modal) return false;
    modal.classList.add('active');
    document.getElementById('modalTitle').textContent = 'Edit: ' + sat.Title;
    setTimeout(() => {
        clearValidationErrors();
        formHandler.setFormData(sat);
    }, 50);
    return false;
}

function deleteSatellite(id) {
    if (!confirm('Delete this satellite?')) return false;
    const index = state.satellites?.findIndex(s => s?.ID === id);
    if (index >= 0) {
        state.satellites.splice(index, 1);
        Logger.info('✅ Deleted');
        saveToLocalStorage();
        showAlert('✅ Deleted!', 'success');
        filterAndDisplayData();
        updateStatistics();
    }
    return false;
}

function selectSatellite(id) {
    const sat = state.satellites?.find(s => s?.ID === id);
    if (!sat) return;
    state.selectedSatellite = sat;
    renderSatelliteDetail();
    filterAndDisplayData();
}

function renderSatelliteDetail() {
    const sat = state.selectedSatellite;
    if (!sat) return;
    const sensors = (sat?.Sensor_Names ?? '').split(',').map(s => s.trim()).filter(s => s);
    document.getElementById('detailTitle').textContent = `Satellite: ${sat?.Title ?? ''}`;
    document.getElementById('detailContent').innerHTML = `<div class="detail-view"><div class="detail-row"><div class="detail-label">Name</div><div class="detail-value">${escapeHtml(sat?.Title ?? '')}</div></div><div class="detail-row"><div class="detail-label">NORAD</div><div class="detail-value">${escapeHtml(sat?.NORAD_ID ?? '')}</div></div><div class="detail-row"><div class="detail-label">Status</div><div class="detail-value">${escapeHtml(sat?.Status ?? '')}</div></div></div>`;
    document.getElementById('detailActions').innerHTML = `<button class="btn btn-small btn-secondary" onclick="editSatellite(${sat?.ID}); return false;">Edit</button><button class="btn btn-small btn-secondary" onclick="deleteSatellite(${sat?.ID}); return false;">Delete</button>`;
}

function switchTab(type, filter) { state.currentFilter = filter; filterAndDisplayData(); return false; }
function sortTable(type, column) {
    if (state.sortColumn === column) { state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc'; } else { state.sortColumn = column; state.sortDirection = 'asc'; }
    filterAndDisplayData();
}
function updateStatistics() {
    document.getElementById('statSatellites').textContent = state.satellites?.length ?? 0;
    document.getElementById('statSensors').textContent = state.sensors?.length ?? 0;
    document.getElementById('statOperational').textContent = (state.satellites ?? []).filter(s => s?.Status === 'Operational').length;
}
function escapeHtml(text) { if (!text) return ''; const div = document.createElement('div'); div.textContent = text; return div.innerHTML; }
function showAlert(message, type = 'info') { const container = document.getElementById('alertContainer'); if (!container) return; const alert = document.createElement('div'); alert.className = `alert alert-${type}`; alert.textContent = message; container.appendChild(alert); setTimeout(() => alert.remove(), 5000); }
function showImportModal() { const modal = document.getElementById('importModal'); if (modal) modal.classList.add('active'); return false; }
function closeImportModal() { const modal = document.getElementById('importModal'); if (modal) modal.classList.remove('active'); return false; }
function handleCSVFileSelect(event) { Logger.info('CSV selected'); }
function handleCSVImport(event) { event.preventDefault(); return false; }
function exportToCSV() { Logger.info('Exporting'); }

// Export formHandler for global access
window.formHandler = formHandler;
console.log('✅ Satellite Sensor Survey v5.5.1 loaded and ready');
    </script>
</body>
</html>
