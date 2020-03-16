import {AddinService} from '../../../controller/AddinService';

export interface IOutlook2SharePointState {
  addinservice: AddinService;
  showSuccess: boolean;
  showError: boolean;
  successMessage: string;
  errorMessage: string;
  isCompose?:boolean;
  cats?:any[];
}
