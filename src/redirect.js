import {broadcastResponseToMainFrame} from '@azure/msal-browser/redirect-bridge';
broadcastResponseToMainFrame().catch(()=>{document.getElementById('status').textContent='Return to the workspace and try signing in again.';});
