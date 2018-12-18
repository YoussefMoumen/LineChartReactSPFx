import * as React from 'react';
import styles from './SuiviFinance.module.scss';
import { ISuiviFinanceProps } from './ISuiviFinanceProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {DetailsListCompactExample} from './Lists';
import LineChart from './lineChart';
import pnp from "@pnp/pnpjs";
import { number } from '@amcharts/amcharts4/core';
export default class SuiviFinance extends React.Component<ISuiviFinanceProps, any> {
public constructor(props: ISuiviFinanceProps){
  super(props);
  this.state = {
    listitems : [{
      key: 0,
      name: "",
      value: 0
    }]
  };
}
public componentDidMount(){
  let _items:[{}];
  pnp.sp.web.lists.getByTitle("List To test").items.get().then(items => 
        items.map((x, key) =>{
          _items.push( { key: key, name:x.Title, value: x.Id });
        }));
        this.setState({listitems: _items});
  }

  public render(): React.ReactElement<ISuiviFinanceProps> {
    return (
      <div>      
      {<DetailsListCompactExample itemes={this.state.listitems}/>}
      {<LineChart />}
      </div>
    );
  }
}