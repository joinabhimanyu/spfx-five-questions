import { IInformaticaFiveQuestionsField } from "./IInformaticaFiveQuestionsState";

export default class MockHttpClient {

  private static lists: Array<any> = [
    {
      key: "05a230f7-a1bd-4f5e-accc-8a3bfbe1dac9",
      text: "Five Questions"
    }
  ];
  private static fields: IInformaticaFiveQuestionsField[] = [
    {
      Id: "05a230f7-a1bd-4f5e-accc-8a3bfbe1dac9",
      ReadOnlyField: false,
      Required: true,
      StaticName: "FieldOne",
      DisplayName: "Field One",
      ValidationFormula: null,
      ValidationMessage: "",
      Value: ""
    },
    {
      Id: "05a230f7-a1bd-4f5e-accc-8a3bfbe1dac8",
      ReadOnlyField: false,
      Required: true,
      StaticName: "FieldTwo",
      DisplayName: "Field Two",
      ValidationFormula: null,
      ValidationMessage: "",
      Value: ""
    },
    {
      Id: "05a230f7-a1bd-4f5e-accc-8a3bfbe1dac7",
      ReadOnlyField: false,
      Required: true,
      StaticName: "FieldThree",
      DisplayName: "Field Three",
      ValidationFormula: null,
      ValidationMessage: "",
      Value: ""
    }
  ];
  public static getLists(): Promise<Array<any>> {
    return new Promise<Array<any>>((resolve) => {
      resolve(MockHttpClient.lists);
    });
  }
  public static getMockFields(): Promise<IInformaticaFiveQuestionsField[]> {
    return new Promise<IInformaticaFiveQuestionsField[]>((resolve) => {
      resolve(MockHttpClient.fields);
    });
  }
}
