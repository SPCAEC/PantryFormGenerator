/** 
 * SPCA Pet Pantry — Generator API Endpoint
 * Exposes controlled access to generator features (e.g., recreate by FormID).
 * Secure via shared SERVICE_API_KEY property.
 */
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) throw new Error('Missing POST data.');
    const payload = JSON.parse(e.postData.contents);
    const { action, formIds, serviceKey } = payload || {};

    const storedKey = PropertiesService.getScriptProperties().getProperty('SERVICE_API_KEY');
    if (!storedKey || serviceKey !== storedKey) throw new Error('Unauthorized');

    switch (action) {
      case 'recreate': {
        const result = recreateFormsById(formIds);
        return jsonOut(result);
      }
      case 'ping':
        return jsonOut({ ok: true, message: 'Generator API online' });
      default:
        throw new Error('Unknown or missing action');
    }
  } catch (err) {
    return jsonOut({ ok: false, message: String(err) });
  }
}

/** Utility: return JSON output */
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}