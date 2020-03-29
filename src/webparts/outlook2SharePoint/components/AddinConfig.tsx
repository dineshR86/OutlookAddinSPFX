import * as React from 'react';
import { AddinService } from '../../../controller/AddinService';
import { PrimaryButton} from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react/lib/Dropdown';

export interface IAddinConfigProps {
  spservice: AddinService;
  stats?: any[];
  cases?: any[];
  configchange?: any;

}

export interface IAddinConfigState {
  isError: boolean;
  errormessage: string;
  localcasses: any[];
  defStatus: string;
  defCase: number;
}

export class AddinConfig extends React.Component<IAddinConfigProps, IAddinConfigState> {
  private _caseid: string = "";
  private _statusid: string = "";


  constructor(props) {
    super(props);
    this.state = {
      isError: false,
      errormessage: "",
      localcasses: [],
      defStatus: "Igangværende",
      defCase: -1
    };
  }

  public componentDidMount() {
    const { spservice } = this.props;
    const configobj = spservice._defConfigData;
    if (typeof configobj != "undefined") {
      spservice.getCases(configobj.Status).then((cases) => {
        this.setState({ localcasses: cases, defCase: Number(configobj.Case), defStatus: configobj.Status });
        this._statusid=configobj.Status;
        this._caseid=configobj.Case;
      });
    }
  }

  public render(): React.ReactElement<IAddinConfigProps> {

    const { isError, errormessage, localcasses, defCase, defStatus } = this.state;
    console.log("Add in Config");

    const casoptions: IDropdownOption[] = this.props != undefined ? this.props.cases : [];
    const statoptions: IDropdownOption[] = this.props != undefined ? this.props.stats : [];
    return (
      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
        <div>
          <Dropdown placeholder="Select an option" label="Configuration Status" options={statoptions} responsiveMode={ResponsiveMode.large} selectedKey={defStatus} onChange={this._statusChange} />
        </div>
        <div style={{ marginTop: "5px" }}>
          <Dropdown placeholder="Select an option" label="Configuration Sager" options={localcasses.length > 0 ? localcasses : casoptions} responsiveMode={ResponsiveMode.large} selectedKey={defCase} onChange={this._casechange} />
        </div>
        <PrimaryButton text="Gem Configuration" onClick={this._onSaveClick} style={{ marginTop: "10px" }} />
      </div>
    );
  }

  private _statusChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    if (option.key != "-1") {
      const { spservice } = this.props;
      this._statusid = option.key.toString();
      this.setState({defStatus:option.key.toString()});
      spservice.getCases(option.key.toString()).then((cases) => {
        this.setState({ localcasses: cases });
      });
    }

  }

  private _casechange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    if (option.key != "-1") {
      this._caseid = option.key.toString();
      this.setState({defCase:Number(option.key)});
    }
  }

  private _onSaveClick = () => {
    if (this._statusid.length <= 0) {
      this._statusid = "Igangværende";
    }
    const configdat = {
      case: this._caseid,
      status: this._statusid
    };

    const { spservice } = this.props;
    const defconfigobj = spservice._defConfigData;
    if (typeof defconfigobj != "undefined") {
      spservice.updateConfigData(configdat,defconfigobj.ID).then((res)=>{
        console.log(res);
        this.props.configchange();
      });

    } else {
      spservice.saveConfigData(configdat).then((dat) => {
        console.log(dat);
        this.props.configchange();
      });
    }
  }
}