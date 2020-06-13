import * as React from 'react';
import styles from './PnPGraphUsers.module.scss';
import { IPnPGraphUsersProps } from './IPnPGraphUsersProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPnPGraphUsersState } from "./IPnPGraphUsersState";
export default class PnPGraphUsers extends React.Component<IPnPGraphUsersProps, IPnPGraphUsersState> {

constructor(props){
  super(props);

  this.state = {
    loading:true
  };
}

 public componentDidMount(){
  this._getMe();
 }

 public _getMe(){
    this.setState({
      loading:true
    },async()=>{
      await this.props.PnPGraphServiceInstance.getCurrentUser();
    })
 }


  public render(): React.ReactElement<IPnPGraphUsersProps> {
    return (
      <div className={ styles.pnPGraphUsers }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
