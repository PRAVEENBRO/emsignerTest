import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/emsignerTestTab/index.html")
@PreventIframe("/emsignerTestTab/config.html")
@PreventIframe("/emsignerTestTab/remove.html")
export class EmsignerTestTab {
}
