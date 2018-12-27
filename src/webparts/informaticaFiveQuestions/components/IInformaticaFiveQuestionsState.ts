export interface IInformaticaFiveQuestionsState {
  validationMessage: string;
  showValidationMsg: boolean;
  validProps: boolean;
  modalVisible: boolean;
  readOnlyModal: boolean;
  ID: number;
  title: string;
  imageFileName: string;
  imageUrl: string;
  authorName: string;
}

export interface IInformaticaFiveQuestionsField {
  Id: string;
  ReadOnlyField: boolean;
  Required: boolean;
  StaticName: string;
  DisplayName: string;
  ValidationFormula: string;
  ValidationMessage: string;
  Value: string;
}

export interface IInformaticaFiveQuestionsFormData {
  fields: IInformaticaFiveQuestionsField[];
  attachement: any;
}
