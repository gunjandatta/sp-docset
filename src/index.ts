import { Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";

/**
 * Global Variable
 */
window["DocSetDemo"] = {
    Configuration
}

// Notify SharePoint that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("docset-demo");