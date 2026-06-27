import { Port } from '@dolphin/common/message'

export enum Flag {
  ExecuteViewScript = 'view_docx_as_markdown',
  ExecuteCopyScript = 'copy_docx_as_markdown',
  ExecuteDownloadScript = 'download_docx_as_markdown',
}

export interface ExecuteScriptMessage {
  flag: Flag
}

export enum RuntimeMessageType {
  FetchAsset = 'fetch_asset',
}

export enum WindowMessageType {
  FetchAssetRequest = 'cdc_fetch_asset_request',
  FetchAssetResponse = 'cdc_fetch_asset_response',
}

export interface FetchAssetMessage {
  type: RuntimeMessageType.FetchAsset
  src: string
}

export interface FetchAssetSuccessResponse {
  ok: true
  dataUrl: string
  contentType: string | null
}

export interface FetchAssetFailureResponse {
  ok: false
  error: string
}

export type FetchAssetResponse =
  | FetchAssetSuccessResponse
  | FetchAssetFailureResponse

export interface WindowFetchAssetRequest {
  type: WindowMessageType.FetchAssetRequest
  id: string
  src: string
}

export interface WindowFetchAssetResponse {
  type: WindowMessageType.FetchAssetResponse
  id: string
  response: FetchAssetResponse
}

export type Message = ExecuteScriptMessage | FetchAssetMessage

export enum EventName {
  Console = 'console',
  GetSettings = 'get_settings',
}

export interface Events extends Record<string, unknown> {
  [EventName.Console]: unknown[]
  [EventName.GetSettings]: string[]
}

class PortImpl {
  private _sender: Port<Events> | null = null
  private _receiver: Port<Events> | null = null

  get sender(): Port<Events> {
    this._sender ??= new Port<Events>('sender', 'receiver')
    return this._sender
  }

  get receiver(): Port<Events> {
    this._receiver ??= new Port<Events>('receiver', 'sender')
    return this._receiver
  }
}

export const portImpl: PortImpl = /* @__PURE__ */ new PortImpl()
