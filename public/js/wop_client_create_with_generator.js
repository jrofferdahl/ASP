/**
 * WOP Client Create with Generator
 * 
 * Client-side wrapper that requests a generated WorkOrderID from the Apps Script
 * (or uses local generateWorkOrderId) and then calls coreGas('createWorkOrder', payload)
 * with WorkOrderID set.
 * 
 * Fallback: if generator is unavailable, proceed without WorkOrderID so server/DB can set it.
 */

(function() {
  'use strict';

  // Configuration - Update with your Apps Script Web App URL after deployment
  window.APPS_SCRIPT_URL = window.APPS_SCRIPT_URL || 'YOUR_WEB_APP_URL_HERE';

  /**
   * Generate Work Order ID locally in format WO-YYYYMMDD-HHMMSS (UTC)
   * @returns {string} Work Order ID
   */
  function generateWorkOrderIdLocal() {
    const now = new Date();
    const year = now.getUTCFullYear();
    const month = String(now.getUTCMonth() + 1).padStart(2, '0');
    const day = String(now.getUTCDate()).padStart(2, '0');
    const hours = String(now.getUTCHours()).padStart(2, '0');
    const minutes = String(now.getUTCMinutes()).padStart(2, '0');
    const seconds = String(now.getUTCSeconds()).padStart(2, '0');
    
    return `WO-${year}${month}${day}-${hours}${minutes}${seconds}`;
  }

  /**
   * Request Work Order ID from Apps Script
   * @returns {Promise<string>} Work Order ID
   */
  async function generateWorkOrderIdFromServer() {
    try {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        return new Promise((resolve, reject) => {
          google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(reject)
            .generateWorkOrderId();
        });
      }
      
      // Fallback to Web App endpoint
      const response = await fetch(window.APPS_SCRIPT_URL + '?op=generateWorkOrderId');
      const data = await response.json();
      
      if (data.ok && data.workOrderId) {
        return data.workOrderId;
      }
      
      throw new Error('Failed to generate Work Order ID from server');
    } catch (error) {
      console.warn('Server generator unavailable, using local:', error);
      return null;
    }
  }

  /**
   * Create work order with generated WorkOrderID
   * @param {Object} options - Options for work order creation
   * @param {string} [options.workOrderText] - Work order description
   * @param {boolean} [options.useServerGenerator=true] - Whether to try server generator first
   * @param {Object} [options.additionalData] - Additional data to include
   * @returns {Promise<Object>} Creation result
   */
  async function createWorkOrderWithGenerator(options = {}) {
    const {
      workOrderText = ' FROM WOP',
      useServerGenerator = true,
      additionalData = {}
    } = options;

    let workOrderId = null;

    // Try to get WorkOrderID from server if enabled
    if (useServerGenerator) {
      workOrderId = await generateWorkOrderIdFromServer();
    }

    // Fallback to local generator
    if (!workOrderId) {
      if (typeof window.generateWorkOrderId === 'function') {
        workOrderId = window.generateWorkOrderId();
      } else {
        workOrderId = generateWorkOrderIdLocal();
      }
    }

    // Build payload
    const payload = {
      WorkOrderID: workOrderId,
      WorkOrderText: workOrderText,
      ...additionalData
    };

    try {
      // Try coreGas if available
      if (typeof coreGas === 'function') {
        return await coreGas('createWorkOrder', payload);
      }
      
      // Try google.script.run
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        return new Promise((resolve, reject) => {
          google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(reject)
            .createWorkOrder(payload);
        });
      }
      
      // Fallback to createNewWorkOrder if available
      if (typeof window.createNewWorkOrder === 'function') {
        return await window.createNewWorkOrder(payload);
      }
      
      throw new Error('No RPC method available for creating work order');
    } catch (error) {
      console.error('Error creating work order with generator:', error);
      
      // Last resort: try without WorkOrderID and let server/DB set it
      if (workOrderId) {
        console.warn('Retrying without WorkOrderID...');
        delete payload.WorkOrderID;
        
        if (typeof window.createNewWorkOrder === 'function') {
          return await window.createNewWorkOrder(payload);
        }
      }
      
      throw error;
    }
  }

  /**
   * Validate Work Order ID format
   * @param {string} workOrderId - Work Order ID to validate
   * @returns {boolean} True if valid format
   */
  function validateWorkOrderIdFormat(workOrderId) {
    if (typeof workOrderId !== 'string') return false;
    
    // Format: WO-YYYYMMDD-HHMMSS
    const pattern = /^WO-\d{8}-\d{6}$/;
    return pattern.test(workOrderId);
  }

  /**
   * Extract date from Work Order ID
   * @param {string} workOrderId - Work Order ID
   * @returns {Date|null} Date object or null if invalid
   */
  function extractDateFromWorkOrderId(workOrderId) {
    if (!validateWorkOrderIdFormat(workOrderId)) return null;
    
    try {
      // Format: WO-YYYYMMDD-HHMMSS
      const parts = workOrderId.split('-');
      const datePart = parts[1]; // YYYYMMDD
      const timePart = parts[2]; // HHMMSS
      
      const year = parseInt(datePart.substring(0, 4), 10);
      const month = parseInt(datePart.substring(4, 6), 10) - 1; // JS months are 0-indexed
      const day = parseInt(datePart.substring(6, 8), 10);
      const hours = parseInt(timePart.substring(0, 2), 10);
      const minutes = parseInt(timePart.substring(2, 4), 10);
      const seconds = parseInt(timePart.substring(4, 6), 10);
      
      return new Date(Date.UTC(year, month, day, hours, minutes, seconds));
    } catch (error) {
      console.error('Error extracting date from Work Order ID:', error);
      return null;
    }
  }

  // Expose functions globally
  window.WOPClient = {
    generateWorkOrderIdLocal,
    generateWorkOrderIdFromServer,
    createWorkOrderWithGenerator,
    validateWorkOrderIdFormat,
    extractDateFromWorkOrderId
  };

  // Also expose top-level for convenience
  window.createWorkOrderWithGenerator = createWorkOrderWithGenerator;
  window.validateWorkOrderIdFormat = validateWorkOrderIdFormat;
  window.extractDateFromWorkOrderId = extractDateFromWorkOrderId;

  console.log('WOP Client loaded - Work Order ID generator available');
})();
