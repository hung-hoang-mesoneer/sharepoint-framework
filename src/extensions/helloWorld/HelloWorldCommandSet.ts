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
    console.log("sharepoint framework ....");
    switch (event.itemId) {
      case "COMMAND_1":
        let selectedItemUrl = event.selectedRows[0].getValueByName("FileRef");
        const fileName = event.selectedRows[0].getValueByName("FileLeafRef");
        console.log(fileName);

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
            console.log("get file from microsoft ok");
            return response.blob();
          })
          .then((fileBlob: Blob) => {
            console.log("sent file to my app");
            this._uploadToServer(fileBlob, fileName);
          });
        // window.open(
        //   "http://localhost:4201/signature-platform/manage-signing-cases",
        //   "_blank"
        // );
        break;
      case "COMMAND_2":
        // Dialog.prompt(`Clicked test. Enter something to alert:`).then(
        //   (value: string) => {
        //     Dialog.alert(value);
        //   }
        // );
        console.log("COMMAND_2 is clicked");
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _uploadToServer(fileBlob: Blob, documentName: string): void {
    const formData = new FormData();
    formData.append("file", fileBlob, documentName);
    const enterpriseId = "bccf1f01-552e-4672-b81b-0c10927ffd4b";
    const organizationId = "bb7dcb58-d3d5-432f-9a00-af1020097a74";
    const documentBasketId = "e9a3a1a3-aadd-4dff-add4-63c1849a3042";

    fetch(
      "http://localhost:8080/internal/signature-platform/" +
        enterpriseId +
        "/organizations/" +
        organizationId +
        "/document-baskets/" +
        documentBasketId +
        "/documents",
      {
        method: "POST",
        body: formData,
      }
    )
      .then((response) => {
        console.log("Document uploaded successfully");
        // Dialog.alert("Document uploaded successfully!");
      })
      .catch((error) => {
        // Dialog.alert("Error uploading document.");
        console.error("Error:", error);
      });
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible =
        this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  };
}
