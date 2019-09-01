import { Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Dashboard } from "./dashboard";
import "./styles.scss";

/**
 * Global Variable
 */
window["DocSetDemo"] = {
    Configuration,
    Dashboard
}

// Notify SharePoint that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("docset-demo");