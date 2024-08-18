import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "HelloWorldCommandSet";

export default class HelloWorldCommandSet extends BaseListViewCommandSet<IHelloWorldCommandSetProperties> {
  public onInit(): Promise<void> {
    console.log(LOG_SOURCE);
    Log.info(LOG_SOURCE, "Initialized HelloWorldCommandSet");

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_1": {
        const selectedItemUrl = event.selectedRows[0].getValueByName("FileRef");
        const fileName = event.selectedRows[0].getValueByName("FileLeafRef");
        this.context.spHttpClient
          .get(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${selectedItemUrl}')/$value`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/octet-stream",
              },
            }
          )
          .then((response: SPHttpClientResponse) => {
            if (!response.ok) {
              throw new Error(`Failed to fetch file: ${response.statusText}`);
            }
            console.log("get file from Microsoft ok");
            return response.blob();
          })
          .then((fileBlob: Blob) => {
            console.log("sent file to signeer");
            this.signWithSigneer(fileBlob, fileName);
          })
          .catch((error) => {
            console.error(
              "There was an error during the file fetching or signing process:",
              error
            );
          });
        break;
      }
      case "COMMAND_2":
        console.log("do COMMAND_2");
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private signWithSigneer(fileBlob: Blob, documentName: string): void {
    console.log("call signWithSigneer.........");
    const formData = new FormData();
    formData.append("file", fileBlob, documentName);
    formData.append(
      "data",
      new Blob(
        [JSON.stringify({ email: this.context.pageContext.user.email })],
        { type: "application/json" }
      )
    );

    fetch(
      "http://localhost:8080/internal/signature-platform/signing-case/document-baskets/sign-with-signeer",
      {
        method: "POST",
        body: formData,
      }
    )
      .then((response) => {
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
      })
      .then((res) => {
        console.log("Document uploaded successfully");
        window.open(res.location, "_blank");
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      const selectedRows = this.context.listView.selectedRows;
      if (selectedRows?.length === 1) {
        const selectedItem = selectedRows[0];

        const fileName: string = selectedItem.getValueByName("FileLeafRef");
        const fileSize: number = selectedItem.getValueByName("File_x0020_Size");

        const isPdfFile = fileName.toLowerCase().endsWith(".pdf");
        const isFileSizeValid = fileSize <= 5 * 1024 * 1024;

        compareOneCommand.visible = isPdfFile && isFileSizeValid;
      } else {
        compareOneCommand.visible = false;
      }

      // You should call this.raiseOnChage() to update the command bar
      this.raiseOnChange();
    }
  };
}
