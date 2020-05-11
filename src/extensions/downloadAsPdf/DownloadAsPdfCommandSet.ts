import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import
{
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { Url, QueryStringCollection, Path, Strings } from '@michaelcoxon/utilities';

import * as strings from 'DownloadAsPdfCommandSetStrings';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.

export interface IDownloadAsPdfCommandSetProperties
{
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}
 */
const SUPPORTED_FORMATS = [
  'csv',
  'doc', 'docx',
  'odp', 'ods', 'odt',
  'pot', 'potm', 'potx', 'pps', 'ppsx', 'ppsxm', 'ppt', 'pptm', 'pptx',
  'rtf',
  'xls', 'xlsx'
];

const LOG_SOURCE: string = 'DownloadAsPdfCommandSet';

export default class DownloadAsPdfCommandSet extends BaseListViewCommandSet<never> {

  @override
  public onInit(): Promise<void>
  {
    Log.info(LOG_SOURCE, 'Initialized DownloadAsPdfCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void
  {
    const compareOneCommand: Command = this.tryGetCommand('ConvertToPdf');
    if (compareOneCommand)
    {
      // This command should be hidden unless exactly one row is selected.

      let visible = true;

      if (event.selectedRows.length == 0)
      {
        visible = visible && false;
      }

      if (!event.selectedRows.every(r => SUPPORTED_FORMATS.indexOf(Strings.trimStart(Path.getExtension(r.getValueByName("FileLeafRef") as string), '.')) > -1))
      {
        visible = visible && false;
      }

      compareOneCommand.visible = visible;
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void>
  {
    switch (event.itemId)
    {
      case 'ConvertToPdf':
        for (const item of event.selectedRows)  
        {
          const filename = item.getValueByName("FileLeafRef") as string;
          const driveUrl = item.getValueByName(".spItemUrl") as string;
          try
          {
            const response = await this.context.spHttpClient.get(driveUrl, SPHttpClient.configurations.v1);
            const json = await response.json();
            const pdfUrl = new Url(Path.combine(json['@odata.id'], "content"), { format: 'pdf' });
            const pdfResponse = await this.context.spHttpClient.get(pdfUrl.toString(), SPHttpClient.configurations.v1);

            if (pdfResponse.ok)
            {
              //this._saveToDisk(pdfResponse.url, filename);
              this._showFile(await pdfResponse.blob(), Path.getFileNameWithoutExtension(filename) + ".pdf");
            }
            else
            {
              Dialog.alert(`File '${filename}' cannot be converted. Please make sure the document is not corrupt.`);
            }
          } catch (ex)
          {
            Dialog.alert(`File '${filename}' cannot be converted. Please make sure the document is not corrupt.`);
          }
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _saveToDisk(fileURL: string, fileName: string)
  {
    // for non-IE
    if (!window['ActiveXObject'])
    {
      const save = document.createElement('a');
      save.href = fileURL;
      save.target = '_blank';
      save.download = fileName || 'unknown';
      if (navigator.userAgent.toLowerCase().match(/(ipad|iphone|safari)/) && navigator.userAgent.search("Chrome") < 0)
      {
        document.location.assign(save.href);
        // window event not working here
      } else
      {
        const evt = new MouseEvent('click', {
          'view': window,
          'bubbles': true,
          'cancelable': false
        });
        save.dispatchEvent(evt);
        (window.URL || window['webkitURL']).revokeObjectURL(save.href);
      }
    }

    // for IE < 11
    else if (!!window['ActiveXObject'] && document.execCommand)
    {
      var _window = window.open(fileURL, '_blank');
      _window.document.close();
      _window.document.execCommand('SaveAs', true, fileName || fileURL);
      _window.close();
    }
  }

  private _showFile(blob: BlobPart, fileName: string)
  {
    // It is necessary to create a new blob object with mime-type explicitly set
    // otherwise only Chrome works like it should
    var newBlob = new Blob([blob], { type: "application/pdf" });

    // IE doesn't allow using a blob object directly as link href
    // instead it is necessary to use msSaveOrOpenBlob
    if (window.navigator && window.navigator.msSaveOrOpenBlob)
    {
      window.navigator.msSaveOrOpenBlob(newBlob);
      return;
    }

    // For other browsers: 
    // Create a link pointing to the ObjectURL containing the blob.
    const data = window.URL.createObjectURL(newBlob);
    var link = document.createElement('a');
    link.href = data;
    link.download = fileName;
    link.click();

    setTimeout(() =>
    {
      // For Firefox it is necessary to delay revoking the ObjectURL
      window.URL.revokeObjectURL(data);
    }, 100);
  }
}
