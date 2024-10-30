import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

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

const LOG_SOURCE: string = 'HelloWorldCommandSet';

export default class HelloWorldCommandSet extends BaseListViewCommandSet<IHelloWorldCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HelloWorldCommandSet');

    // Initial state of the command's visibility
    const compareOneCommand: Command | undefined = this.tryGetCommand('MARK_CUSTOMER');
    if (compareOneCommand) {
      compareOneCommand.visible = false;
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'MARK_CUSTOMER':
        event.selectedRows.map(row => {
          // Get lead data
          const id: number = row.getValueByName('ID');
          const isCustomer: boolean = row.getValueByName('Customer') === 'Yes';

          const body: string = JSON.stringify({
            'Customer': isCustomer ? 'No' : 'Yes'
          });

          this.context.spHttpClient.post(
            `${this.context.pageContext.web.absoluteUrl}/_api/lists/getbyid('${this.context.pageContext.list?.id}')/items(${id})`,
            SPHttpClient.configurations.v1,
            {
              body: body,
              headers: {
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*'
              }
            }
          ).then((response: SPHttpClientResponse): void => {
            if (response.ok) {
              console.log(`Item with ID: ${id} successfully updated.`);
            } else {
              console.error(`Failed to update item with ID: ${id}.`);
            }
          }).catch(error => {
            console.error(`Error updating item with ID: ${id}. Error: ${error}`);
          });
        });
        break;

      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command | undefined = this.tryGetCommand('MARK_CUSTOMER');
    if (compareOneCommand) {
      // This command should be visible when exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
      this.raiseOnChange(); // Notify the framework to update the command bar
    }
  }
}
