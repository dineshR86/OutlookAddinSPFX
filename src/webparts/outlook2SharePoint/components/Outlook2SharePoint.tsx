import * as React from 'react';
import styles from './Outlook2SharePoint.module.scss';
import { AddinService } from '../../../controller/AddinService';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AddinCompose } from './AddinCompose';
import { AddinRead } from './AddinRead';
import {AddinConfig} from './AddinConfig';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IconButton,IIconProps } from 'office-ui-fabric-react';


const settingsIcon: IIconProps = { iconName: 'PlayerSettings' };

export interface IOutlook2SharePointProps {
  mail: any;
  context: WebPartContext;
  msGraphClientFactory: MSGraphClientFactory;
}

export interface IOutlook2SharePointState {
  addinservice: AddinService;
  isCompose?: boolean;
  cats?: any[];
  stats?: any[];
  cases?: any[];
  loading?: boolean;
  showconfig?:boolean;
}

export default class Outlook2SharePoint extends React.Component<IOutlook2SharePointProps, IOutlook2SharePointState> {
  private _addinservice: AddinService;

  constructor(props) {
    super(props);
    this.state = {
      addinservice: null,
      isCompose: false,
      cats: [],
      stats: [],
      cases: [],
      loading: false,
      showconfig:false
    };

    this._addinservice = new AddinService(this.props.context, this.props.mail, this.props.msGraphClientFactory);
  }

  public componentDidMount() {
    this.setState({loading:true});
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let myParm: boolean = queryParms.getValue("isCompose") == "true" ? true : false;
    this.setState({ isCompose: myParm });

    this._addinservice.getCategories().then((dat) => {
      this.setState({ cats: dat });
    });

    this._addinservice.getCaseStatus().then((sdat) => {
      this.setState({ stats: sdat });
    });

    this._addinservice.getCases("Igangværende").then((cdat) => {
      this.setState({ cases: cdat,loading:false });
    });
  }

  public render(): React.ReactElement<IOutlook2SharePointProps> {
    const { isCompose, cats, stats, cases,loading,showconfig } = this.state;
    const spfxstyles = {
      spinner: {
        display: loading ? "block" : "none"
      }
    };
    
    return (
      <div className="ms-Grid" dir="ltr">
        <div style={spfxstyles.spinner}>
          <Spinner label="Loading the Addin..." />
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm10 ms-md10 ms-lg10"></div>
          <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
          <IconButton iconProps={settingsIcon} title="User Settings" ariaLabel="PlayerSettings" onClick={this._onSettingsClick} />
          </div>
        </div>
       {this._onrender()}
      </div>
    );
  }

  private _onSettingsClick=()=>{
    const{showconfig}=this.state;
    this.setState({showconfig:!showconfig});
  }

  private _onrender(){
    const{isCompose, cats, stats, cases,loading,showconfig}=this.state;
    if(showconfig){
      return <AddinConfig spservice={this._addinservice} stats={stats} cases={cases} configchange={this._configchange} />;
    }else{
      return (
        <div>
        <div className="ms-Grid-row">
          {isCompose ? <AddinCompose spservice={this._addinservice} categories={cats} stats={stats} cases={cases} /> : <AddinRead spservice={this._addinservice} categories={cats} stats={stats} cases={cases} />}
        </div>
        </div>
      );
    }
  }

  private _configchange=()=>{
    this.setState({showconfig:false});
  }
}
