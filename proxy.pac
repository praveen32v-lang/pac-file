function FindProxyForURL(url, host)
{
    var direct = "DIRECT";
    var proxyServer = "PROXY turbo-qhxk42r2.edge.prod.fortisase.com:11508";

    // Use proxy for specific Office CDN domains
    if (shExpMatch(host, "cdn.odc.officeapps.live.com") ||
        shExpMatch(host, "cdn.uci.officeapps.live.com")) {
        return proxyServer;
    }

    // Go direct for Microsoft 365 / Teams / Identity services
    if (shExpMatch(host, "*.auth.microsoft.com") ||
        shExpMatch(host, "*.lync.com") ||
        shExpMatch(host, "*.mail.protection.outlook.com") ||
        shExpMatch(host, "*.msftidentity.com") ||
        shExpMatch(host, "*.msidentity.com") ||
        shExpMatch(host, "*.mx.microsoft") ||
        shExpMatch(host, "*.officeapps.live.com") ||
        shExpMatch(host, "*.online.office.com") ||
        shExpMatch(host, "*.protection.office.com") ||
        shExpMatch(host, "*.protection.outlook.com") ||
        shExpMatch(host, "*.security.microsoft.com") ||
        shExpMatch(host, "*.sharepoint.com") ||
        shExpMatch(host, "*.teams.cloud.microsoft") ||
        shExpMatch(host, "*.teams.microsoft.com") ||
        shExpMatch(host, "account.activedirectory.windowsazure.com") ||
        shExpMatch(host, "accounts.accesscontrol.windows.net") ||
        shExpMatch(host, "adminwebservice.microsoftonline.com") ||
        shExpMatch(host, "api.passwordreset.microsoftonline.com") ||
        shExpMatch(host, "autologon.microsoftazuread-sso.com") ||
        shExpMatch(host, "becws.microsoftonline.com") ||
        shExpMatch(host, "ccs.login.microsoftonline.com") ||
        shExpMatch(host, "clientconfig.microsoftonline-p.net") ||
        shExpMatch(host, "companymanager.microsoftonline.com") ||
        shExpMatch(host, "compliance.microsoft.com") ||
        shExpMatch(host, "defender.microsoft.com") ||
        shExpMatch(host, "device.login.microsoftonline.com") ||
        shExpMatch(host, "graph.microsoft.com") ||
        shExpMatch(host, "graph.windows.net") ||
        shExpMatch(host, "login.microsoft.com") ||
        shExpMatch(host, "login.microsoftonline.com") ||
        shExpMatch(host, "login.microsoftonline-p.com") ||
        shExpMatch(host, "login.windows.net") ||
        shExpMatch(host, "logincert.microsoftonline.com") ||
        shExpMatch(host, "loginex.microsoftonline.com") ||
        shExpMatch(host, "login-us.microsoftonline.com") ||
        shExpMatch(host, "nexus.microsoftonline-p.com") ||
        shExpMatch(host, "office.live.com") ||
        shExpMatch(host, "outlook.cloud.microsoft") ||
        shExpMatch(host, "outlook.office.com") ||
        shExpMatch(host, "outlook.office365.com") ||
        shExpMatch(host, "passwordreset.microsoftonline.com") ||
        shExpMatch(host, "protection.office.com") ||
        shExpMatch(host, "provisioningapi.microsoftonline.com") ||
        shExpMatch(host, "purview.microsoft.com") ||
        shExpMatch(host, "security.microsoft.com") ||
        shExpMatch(host, "smtp.office365.com") ||
        shExpMatch(host, "teams.cloud.microsoft") ||
        shExpMatch(host, "teams.microsoft.com")) {
        return direct;
    }

    // Go direct for Microsoft Diagnostics & Telemetry
    if (shExpMatch(host, "*.diagnostics.office.com") ||
        shExpMatch(host, "*.diagnostics-eudb.office.com") ||
        shExpMatch(host, "incidents.diagnostics-eudb.office.com") ||
        shExpMatch(host, "*.events.data.microsoft.com") ||
        shExpMatch(host, "*.telemetry.microsoft.com")) {
        return direct;
    }

    // Go direct for specific IP exclusions
    if (isInNet(host, "82.218.103.242", "255.255.255.255") ||
        isInNet(host, "80.120.163.34", "255.255.255.255") ||
        isInNet(host, "187.217.233.162", "255.255.255.255") ||
        isInNet(host, "194.228.30.74", "255.255.255.255") ||
        isInNet(host, "61.161.218.210", "255.255.255.255") ||
        isInNet(host, "176.94.250.46", "255.255.255.255") ||
        isInNet(host, "173.167.30.1", "255.255.255.255") ||
        isInNet(host, "91.192.2.168", "255.255.255.255") ||
        isInNet(host, "95.65.10.74", "255.255.255.255") ||
        isInNet(host, "178.239.68.82", "255.255.255.255")) {
        return direct;
    }

    // Default: use proxy
    return proxyServer;
}


