import * as React from 'react';
import { AddinService } from '../../../controller/AddinService';
import styles from './Outlook2SharePoint.module.scss';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react/lib/Dropdown';
import { PrimaryButton, MessageBar, MessageBarType } from 'office-ui-fabric-react';

export interface IAddinComposeProps {
  spservice: AddinService;
  categories?: any[];
  stats?: any[];
  cases?: any[];
}

export interface IAddinComposeState {
  isCatVisible: boolean;
  casechange: boolean;
  isError: boolean;
  errormessage: string;
  localcases?: any[];
  defStatus: string;
  defCase: number;
}

export class AddinCompose extends React.Component<IAddinComposeProps, IAddinComposeState> {
  private _caseid: string = "";
  private _catid: string = "";

  constructor(props) {
    super(props);
    this.state = {
      isCatVisible: false,
      casechange: false,
      isError: false,
      errormessage: "",
      localcases: [],
      defStatus: "Igangværende",
      defCase: -1
    };
  }

  public componentDidMount() {
    const { spservice } = this.props;
    spservice._mail.subject.getAsync((result) => {
      spservice._mailsubject = result.value;
    });
    spservice.getConfigData().then((configObj) => {
      if (typeof configObj != "undefined") {
        spservice._defConfigData = configObj;
        this.setState({ defCase: Number(configObj.Case), defStatus: configObj.Status });
        spservice.getCases(configObj.Status).then((cases) => {
          this.setState({ localcases: cases, casechange: false });
        });
        this.setState({ isCatVisible: true });
        this._caseid = configObj.Case.toString();
      }
    });

    // spservice._mail.body.getAsync('text', (result)=> {
    //   if (result.status === 'succeeded') {
    //     spservice._mailbody=result.value;
    //   }
    // });
  }

  public render(): React.ReactElement<IAddinComposeProps> {

    const { isCatVisible, isError, errormessage, localcases,defCase,defStatus } = this.state;

    const catoptions: IDropdownOption[] = this.props != undefined ? this.props.categories : [];
    const casoptions: IDropdownOption[] = this.props != undefined ? this.props.cases : [];
    const statoptions: IDropdownOption[] = this.props != undefined ? this.props.stats : [];

    const addinstyles = {
      catvisible: {
        display: isCatVisible ? "block" : "none",
        marginBottom: "20px",
      },
      errormessage: {
        display: isError ? "block" : "none"
      }
    };

    const options: IDropdownOption[] = this.props != undefined ? this.props.categories : [];
    return (
      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
        <div style={addinstyles.errormessage}>
          <MessageBar messageBarType={MessageBarType.error} >
            {errormessage}
          </MessageBar>
        </div>
        <div>
          <Dropdown placeholder="Select an option" label="Status" options={statoptions} responsiveMode={ResponsiveMode.large} selectedKey={defStatus} onChange={this._statusChange} />
        </div>
        <div style={{ marginTop: "5px" }}>
          <Dropdown placeholder="Select an option" label="Sager" options={localcases.length > 0 ? localcases : casoptions} responsiveMode={ResponsiveMode.large} selectedKey={defCase} onChange={this._casechange} />
        </div>
        <div style={addinstyles.catvisible}>
          <Dropdown placeholder="Select an option" label="Kategori" options={catoptions} responsiveMode={ResponsiveMode.large} onChange={this._catchange} />
          <PrimaryButton text="Gem" onClick={this._onSaveClick} style={{ marginTop: "20px" }} />
        </div>
      </div>
    );
  }

  private _statusChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    if (option.key != "-1") {
      const { spservice } = this.props;
      spservice.getCases(option.key.toString()).then((cases) => {
        this.setState({ localcases: cases, casechange: false });
      });
    }

  }

  private _onSaveClick = () => {
    if (this._catid != "-1" && this._catid.length > 0) {
      this.props.spservice.composemail(`ID${this._caseid}- Cat${this._catid}`);
    }
    else {
      this.setState({ errormessage: "Please select category or subject is missing", isError: true });
    }
  }

  private _casechange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._caseid = option.key.toString();
    if (option.key.toString() != "-1") {
      this.setState({ isCatVisible: true });
    } else {
      this.setState({ isCatVisible: false });
    }
  }

  private _catchange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._catid = option.key.toString();
  }

}