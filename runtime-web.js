/*
 * Outlook on the web, new Outlook for Windows, and Outlook for Mac load the
 * event runtime through runtime.html. Calling Office.onReady initializes that
 * host so the actions registered in runtime.js can receive launch events.
 */
Office.onReady();
