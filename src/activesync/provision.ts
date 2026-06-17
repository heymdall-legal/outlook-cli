import type { BaseExchangeConfig } from "../types.js";
import { getFirstTagText } from "./xml.js";

const POLICY_TYPE = "MS-EAS-Provisioning-WBXML";

interface ProvisionRequestConfig
  extends Pick<
    BaseExchangeConfig,
    | "deviceModel"
    | "deviceImei"
    | "deviceFriendlyName"
    | "deviceOs"
    | "deviceOsLanguage"
    | "devicePhoneNumber"
    | "deviceMobileOperator"
    | "userAgent"
  > {}

export interface ProvisionResponse {
  status: string;
  policyType: string;
  policyStatus: string;
  policyKey: string;
  remoteWipe: boolean;
}

export function buildInitialProvisionRequestXml(config: ProvisionRequestConfig): string {
  return `<?xml version="1.0" encoding="utf-8"?><Provision xmlns="Provision:" xmlns:settings="Settings:"><settings:DeviceInformation><settings:Set><settings:Model>${escapeXml(
    config.deviceModel
  )}</settings:Model><settings:IMEI>${escapeXml(config.deviceImei)}</settings:IMEI><settings:FriendlyName>${escapeXml(
    config.deviceFriendlyName
  )}</settings:FriendlyName><settings:OS>${escapeXml(config.deviceOs)}</settings:OS><settings:OSLanguage>${escapeXml(
    config.deviceOsLanguage
  )}</settings:OSLanguage><settings:PhoneNumber>${escapeXml(
    config.devicePhoneNumber
  )}</settings:PhoneNumber><settings:MobileOperator>${escapeXml(
    config.deviceMobileOperator
  )}</settings:MobileOperator><settings:UserAgent>${escapeXml(
    config.userAgent
  )}</settings:UserAgent></settings:Set></settings:DeviceInformation><Policies><Policy><PolicyType>${POLICY_TYPE}</PolicyType></Policy></Policies></Provision>`;
}

export function buildProvisionAckRequestXml({
  policyKey,
  status = "1",
}: {
  policyKey: string;
  status?: string;
}): string {
  return `<?xml version="1.0" encoding="utf-8"?><Provision xmlns="Provision:"><Policies><Policy><PolicyType>${POLICY_TYPE}</PolicyType><PolicyKey>${escapeXml(
    policyKey
  )}</PolicyKey><Status>${escapeXml(status)}</Status></Policy></Policies></Provision>`;
}

export function parseProvisionResponseXml(xml: string): ProvisionResponse {
  const policyBlock = xml.match(/<Policy(?:\s[^>]*)?>([\s\S]*?)<\/Policy>/i)?.[1] ?? "";

  return {
    status: getFirstTagText(xml, "Status"),
    policyType: getFirstTagText(policyBlock, "PolicyType"),
    policyStatus: getFirstTagText(policyBlock, "Status"),
    policyKey: getFirstTagText(policyBlock, "PolicyKey"),
    remoteWipe: xml.includes("<RemoteWipe"),
  };
}

function escapeXml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
