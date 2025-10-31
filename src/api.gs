/**
 * SPCA Pet Pantry — Generator API Endpoint
 * Exposes controlled access to generator features (e.g., recreate by FormID).
 * Secure via shared SERVICE_API_KEY property.
 */
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) throw new Error('Missing POST data.');
    const payload = JSON.parse(e.postData.contents || '{}');
    const { action, formIds, serviceKey } = payload;

    const storedKey = PropertiesService.getScriptProperties().getProperty('SERVICE_API_KEY');
    if (!storedKey || serviceKey !== storedKey) throw new Error('Unauthorized');

    switch (action) {
      case 'recreate':
        return jsonOut(handleRecreateRequest_(formIds));

      case 'ping':
        return jsonOut({ ok: true, message: 'Generator API online' });

      default:
        throw new Error('Unknown or missing action');
    }
  } catch (err) {
    Logger.log('❌ doPost error: %s', err.stack || err);
    return jsonOut({ ok: false, message: String(err) });
  }
}

/**
 * Handle incoming recreate requests from the Daily Pantry Dashboard.
 * Accepts one or more FormIDs and regenerates each via internal generator logic.
 */
function handleRecreateRequest_(formIds) {
  try {
    if (!formIds || !formIds.length) throw new Error('No FormIDs provided.');
    if (!Array.isArray(formIds)) formIds = [formIds];

    const results = [];
    formIds.forEach(id => {
      try {
        // Core function: regenerate this form by FormID
        const output = generateOrderFormByFormId(id, true);
        results.push({ formId: id, ok: true, output });
      } catch (errInner) {
        Logger.log('⚠️ Recreate failed for %s: %s', id, errInner);
        results.push({ formId: id, ok: false, message: String(errInner) });
      }
    });

    const okCount = results.filter(r => r.ok).length;
    return {
      ok: true,
      count: okCount,
      results,
      message: `Recreated ${okCount} of ${results.length} form(s) successfully.`
    };
  } catch (err) {
    Logger.log('❌ handleRecreateRequest_ error: %s', err.stack || err);
    return { ok: false, message: String(err) };
  }
}

/** Utility: return JSON output */
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}