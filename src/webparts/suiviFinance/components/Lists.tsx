import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import pnp from "@pnp/pnpjs";
const _items: {
  key: number;
  name: string;
  value: number;
}[] = [];

const _columns = [
  {
    key: 'column1',
    name: 'Name',
    fieldName: 'name',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column2',
    name: 'Value',
    fieldName: 'value',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  }
];

export class DetailsListCompactExample extends React.Component<
  {itemes:any},
  {
    items: {}[];
    selectionDetails: string;
  }
> {
  private _selection: Selection;

  constructor(props: {itemes:any}) {
    super(props);
    let {itemes} = this.props;
    console.log("Properties In List Component=>",this.props);
    // Populate with items for demos.
    // if (_items.length === 0) {
    //   pnp.sp.web.lists.getByTitle("List To test").items.get().then(items => 
    //     items.map((x, key) =>{
    //       _items.push( { key: key, name:x.Title, value: x.Id });
    //     }));
    //     console.log("_items =>", _items);
    // }

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
    });

    this.state = {
      // items: _items,
      items: itemes,
      selectionDetails: this._getSelectionDetails()
    };
  }
  public getListItemsByTitle(listTitle: string):Promise<any>{
    return pnp.sp.web.lists.getByTitle(listTitle).items.get().then(items => 
      items.map((x, key) =>{
        _items.push( { key: key, name:x.Title, value: x.Id });
      }));
    }

  public render(): JSX.Element {
    const { items, selectionDetails } = this.state;

    return (
      <div>
        <div>{selectionDetails}</div>
        <TextField label="Filter by name:" />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            columns={_columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            compact={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
        </MarqueeSelection>
      </div>
    );
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as any).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({ items: text ? _items.filter(i => i.name.toLowerCase().indexOf(text) > -1) : _items });
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }
}