import InformaticaFiveQuestions from "./InformaticaFiveQuestions";
import { IInformaticaFiveQuestionsFormData } from "./IInformaticaFiveQuestionsState";
import { connect } from "react-redux";
import {
  setFormData,
  resetFormData,
  setErrMsg
} from "./redux";

const mapStateToProps: any = (state, ownProps) => ({
  formData: state.formData,
  errorMessage: state.errorMessage,
  WebPartTitle: ownProps.WebPartTitle,
  context: ownProps.context,
  ListName: ownProps.ListName
});

const mapDispatchToProps: any = {
  setFormData,
  resetFormData,
  setErrMsg
};

const AppContainer: any = connect(
  mapStateToProps,
  mapDispatchToProps
)(InformaticaFiveQuestions);

export default AppContainer;
