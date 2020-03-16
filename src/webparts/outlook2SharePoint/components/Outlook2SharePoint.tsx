import * as React from 'react';
import styles from './Outlook2SharePoint.module.scss';
import { AddinService } from '../../../controller/AddinService';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AddinCompose } from './AddinCompose';
import { AddinRead } from './AddinRead';
import { MSGraphClientFactory } from '@microsoft/sp-http';

export interface IOutlook2SharePointProps {
  mail: any;
  context:WebPartContext;
  msGraphClientFactory: MSGraphClientFactory;
}

export interface IOutlook2SharePointState {
  addinservice: AddinService;
  isCompose?:boolean;
  cats?:any[];
  stats?:any[];
  cases?:any[];
}

export default class Outlook2SharePoint extends React.Component<IOutlook2SharePointProps, IOutlook2SharePointState> {
  private _addinservice: AddinService;

  constructor(props) {
    super(props);
    this.state = {
      addinservice: null,
      isCompose: false,
      cats:[],
      stats:[],
      cases:[]
    };

    this._addinservice = new AddinService(this.props.context,this.props.mail,this.props.msGraphClientFactory);
  }

  public componentDidMount() {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let myParm: boolean = queryParms.getValue("isCompose") == "true" ? true : false;
    this.setState({ isCompose: myParm });

    this._addinservice.getCategories().then((dat) => {
      this.setState({cats:dat});
    });

    this._addinservice.getCaseStatus().then((sdat)=>{
      this.setState({stats:sdat});
    });

    this._addinservice.getCases("Igangværende").then((cdat)=>{
      this.setState({cases:cdat});
    });
  }

  public render(): React.ReactElement<IOutlook2SharePointProps> {
    const { isCompose,cats,stats,cases } = this.state;
    return (
      <div className="ms-Grid" dir="ltr">
        <div className="ms-Grid-row">
          {isCompose ? <AddinCompose spservice={this._addinservice} categories={cats} stats={stats} cases={cases} /> : <AddinRead spservice={this._addinservice} categories={cats} stats={stats} cases={cases} />}
        </div>
      </div>
    );
  }
}
