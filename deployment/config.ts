import { getFlight } from "../utils/flights";
import { getWidgetTheme } from "./utils/getWidgetTheme";
import { handleImplicitData } from "./utils/handleImplicitData";
import { WidgetDoneCallback } from "../types/WidgetDoneCallback";
import { showFlowWidget } from "./utils/showWidget";
import { getPowerAutomateAccessToken, isTokenExpired } from "./utils/tokenUtils";

declare var MsFlowSdk: any;

export function loadFlowWidget(endpoint: string, exelHostDriveid: string, excelFileId: string, autoFillParams: any, flowToken: string) {
    let sdk = new MsFlowSdk({
        hostName: endpoint,
        locale: Office.context.displayLanguage || 'en-us',
        hostId: 'ExcelSDX',
        hostLocale: Office.context.displayLanguage,
        hostVersion: Office.context.diagnostics.version,
        hostPlatform: Office.context.platform.toString(),
        enableWidgetV2: true,
    });

    const widgetRenderParams = {
        debugMode: getFlight("PowerAutomatePPUXDevMode"),
        container: 'flow-div',
        flowsSettings: {
            allowImplicitConsent: true,
            hideTabs: true,
            isMini: true,
            flowsFilter: `operations/any(operation: operation/excel.fileId eq '${exelHostDriveid}/${excelFileId}')`,
            widgetFlowListDisplaySettings: {
                actionMenuOverFlowItems: true,
                actionMenuClassName: 'fl-ActionMenu-ExcelWidget',
                triggerOperationKey: 'SHARED_EXCELONLINEBUSINESS-ONROWSELECTED',
                triggerOperationName: 'OnRowSelected',
                triggerOperationGroupName: 'shared_excelonlinebusiness',
                hideTemplateTitleDietDesigner: true,
                hideTemplateTypeDietDesigner: true,
            },
        },
        templatesSettings: {
            allowCustomFlowName: true,
            metadataSortProperty: 'ExcelTablesPriority',
            templateCategory: 'microsoftexcel_sdx_nativem2',
            useFlowCreatorSurfaceFromTemplateGallery: true,
            enableDietDesigner: true,
            showHiddenTemplates: true,
            enableTemplatesPageShell: true,
            defaultParams: autoFillParams,
        },
        enableOnBehalfOfTokens: true,
        widgetStyleSettings: {
            themeName: getWidgetTheme(),
        }
    };

    let widgetInstance = sdk.renderWidget('flows', widgetRenderParams);

    widgetInstance.listen('GET_ACCESS_TOKEN', async (_requestParams: any, widgetDoneCallback: WidgetDoneCallback) => {
        // If the current token is expired, get a new one
        if (isTokenExpired(flowToken)) {
            const tokenResponse = await getPowerAutomateAccessToken();
            flowToken = tokenResponse.accessToken;
        }

        widgetDoneCallback(null, { token: flowToken });
    });

    widgetInstance.listen('WIDGET_READY', () => {
        showFlowWidget();
    });

    widgetInstance.listen('GET_IMPLICIT_DATA', (requestParam: { data: { implicitData?: object } }, widgetDoneCallback: WidgetDoneCallback) => {
        handleImplicitData(requestParam.data, widgetDoneCallback);
    });
}
