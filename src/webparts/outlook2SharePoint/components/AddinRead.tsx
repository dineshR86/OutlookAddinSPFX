import * as React from 'react';
import { AddinService } from '../../../controller/AddinService';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { PrimaryButton } from 'office-ui-fabric-react';
import styles from './Outlook2SharePoint.module.scss';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react/lib/Dropdown';

export interface IAddinReadProps {
  spservice: AddinService;
  categories?: any[];
  stats?: any[];
  cases?: any[];
}

export interface IAddinReadState {
  isCatVisible: boolean;
  drop1: any[];
  drop2: any[];
  drop3: any[];
  drop4: any[];
  drop5: any[];
  saveemail: boolean;
  saveattachment: boolean;
  casechange:boolean;
  folderpath?: string;
}

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};



export class AddinRead extends React.Component<IAddinReadProps, IAddinReadState> {
  private _caseTitle: string;
  private _caseid:string;
  private _catid:string;
  private _drop1: string;
  private _drop2: string;
  private _drop3: string;
  private _drop4: string;
  private _drop5: string;

  constructor(props) {
    super(props);
    this.state = {
      isCatVisible: false,
      drop1: [],
      drop2: [],
      drop3: [],
      drop4: [],
      drop5: [],
      saveemail: false,
      saveattachment: false,
      casechange:false
    };
  }

  public render(): React.ReactElement<IAddinReadProps> {
    const { isCatVisible, drop1, drop2, drop3, drop4, drop5,saveattachment,saveemail,casechange } = this.state;
    const addinstyles = {
      catvisible: {
        display: isCatVisible ? "block" : "none"
      },
      drop1: {
        display: drop1.length > 0 ? "block" : "none",
        marginTop: "20px"
      },
      drop2: {
        display: drop2.length > 0 ? "block" : "none",
        marginTop: "20px"
      },
      drop3: {
        display: drop3.length > 0 ? "block" : "none",
        marginTop: "20px"
      },
      drop4: {
        display: drop4.length > 0 ? "block" : "none",
        marginTop: "20px"
      },
      drop5: {
        display: drop5.length > 0 ? "block" : "none",
        marginTop: "20px"
      },
      saveemail:{
        display:saveemail?"block":"none"
      },
      saveattachment:{
        display:saveattachment?"block":"none"
      },
      caseChange:{
        display:casechange?"block":"none"
      }
    };

    const catoptions: IDropdownOption[] = this.props != undefined ? this.props.categories : [];
    const casoptions: IDropdownOption[] = this.props != undefined ? this.props.cases : [];
    const statoptions: IDropdownOption[] = this.props != undefined ? this.props.stats : [];

    return (
      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
        <div>
          <Dropdown placeholder="Select an option" label="Status" options={statoptions} responsiveMode={ResponsiveMode.large} defaultSelectedKey="Igangværende" />
        </div>
        <div>
          <Dropdown placeholder="Select an option" label="Sager" options={casoptions} responsiveMode={ResponsiveMode.large} onChange={this._casechange} />
        </div>
        <div style={addinstyles.caseChange}>
        <div style={{marginTop:"20px"}}>
          <Checkbox label="Gem Email" onChange={this._saveemail} defaultChecked />

          <div style={addinstyles.catvisible}>
            <Dropdown placeholder="Select an option" label="Kategori" options={catoptions} responsiveMode={ResponsiveMode.large} onChange={this._catchange} />
          </div>
        </div>
        <div style={{marginTop:"20px", marginBottom:"20px"}}>
          <Checkbox label="Gem Vedhæftning(er)" onChange={this._saveattachment} />
          <div style={addinstyles.saveattachment}>
          <div style={addinstyles.drop1}>
            <Dropdown placeholder="Select an option" label="" options={drop1} responsiveMode={ResponsiveMode.large} onChange={this._drop1Change} />
          </div>
          <div style={addinstyles.drop2}>
            <Dropdown placeholder="Select an option" label="" options={drop2} responsiveMode={ResponsiveMode.large} onChange={this._drop2Change} />
          </div>
          <div style={addinstyles.drop3}>
            <Dropdown placeholder="Select an option" label="" options={drop3} responsiveMode={ResponsiveMode.large} onChange={this._drop3Change} />
          </div>
          <div style={addinstyles.drop4}>
            <Dropdown placeholder="Select an option" label="" options={drop4} responsiveMode={ResponsiveMode.large} onChange={this._drop4Change} />
          </div>
          <div style={addinstyles.drop5}>
            <Dropdown placeholder="Select an option" label="" options={drop5} responsiveMode={ResponsiveMode.large} />
          </div>
          </div>
        </div>
          <PrimaryButton text="Gem" onClick={this._onSaveClick} />
        </div>
      </div>
    );
  }

  private _casechange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this.setState({ isCatVisible: true });
    const { spservice } = this.props;
    this._caseid=option.key.toString();
    spservice.getCaseFolderTitle(option.key.toString()).then((tit) => {
      spservice.getCaseSubFolders(tit).then((dat) => {
        this.setState({ drop1: dat,casechange:true });
      });
      //this.setState({folderpath:`${tit}`});
      this._caseTitle = tit;
    });
  }

  private _catchange=(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number)=>{
    this._catid=option.key.toString();
  }

  private _drop1Change = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._drop1 = option.key.toString();
    if (this._drop1 != "-1") {
      this.props.spservice.getCaseSubFolders(`${this._caseTitle}/${this._drop1}`).then((dat2) => {
        this.setState({ drop2: dat2 });
      });
    }
  }

  private _drop2Change = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._drop2 = option.key.toString();
    if (this._drop2 != "-1") {
      debugger;
      this.props.spservice.getCaseSubFolders(`${this._caseTitle}/${this._drop1}/${this._drop2}`).then((dat3) => {
        this.setState({ drop3: dat3 });
      });
    }
  }

  private _drop3Change = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._drop3 = option.key.toString();
    if (this._drop3 != "-1") {
      debugger;
      this.props.spservice.getCaseSubFolders(`${this._caseTitle}/${this._drop1}/${this._drop2}/${this._drop3}`).then((dat4) => {
        this.setState({ drop4: dat4 });
      });
    }
  }

  private _drop4Change = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    this._drop4 = option.key.toString();
    if (this._drop4 != "-1") {
      debugger;
      this.props.spservice.getCaseSubFolders(`${this._caseTitle}/${this._drop1}/${this._drop2}/${this._drop3}/${this._drop4}`).then((dat5) => {
        this.setState({ drop5: dat5 });
      });
    }
  }

  private _saveemail = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
    this.setState({saveemail:isChecked});
  }

  private _saveattachment = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
    this.setState({saveattachment:isChecked});
  }

  private _onSaveClick=()=>{
    const addinobj={
      catid:this._catid,
      caseid:this._caseid
    };

    this.props.spservice.saveemail(addinobj).then((dat)=>{
      console.log(dat);
      Office.context.ui.closeContainer();
    });

    //this.props.spservice.getAttachments();
  }

}