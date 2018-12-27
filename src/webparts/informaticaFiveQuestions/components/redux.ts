import {
  applyMiddleware,
  combineReducers,
  createStore,
} from "redux";
import { IInformaticaFiveQuestionsFormData } from "./IInformaticaFiveQuestionsState";

// actions.js
export const setFormData: Function = (formDataArgs: IInformaticaFiveQuestionsFormData) => ({
  type: "SET_FORMDATA",
  payload: { fields: formDataArgs.fields, attachement: formDataArgs.attachement },
});

export const resetFormData: Function = () => ({
  type: "RESET_FORMDATA",
});

export const setErrMsg: any = (errMsg: string) => ({
  type: "SET_ERRMSG",
  payload: errMsg
});

const initialState: IInformaticaFiveQuestionsFormData = {
  fields: [],
  attachement: null
};

// reducers.js
export const formData: any = (state: IInformaticaFiveQuestionsFormData = initialState, action): IInformaticaFiveQuestionsFormData => {
  switch (action.type) {
    case "SET_FORMDATA":
      return { fields: action.payload.fields, attachement: action.payload.attachement };
    case "RESET_FORMDATA":
      return initialState;
    default:
      return state;
  }
};

export const errorMessage: any = (state: string = "", action): string => {
  switch (action.type) {
    case "SET_ERRMSG":
      return action.payload;
    default:
      return state;
  }
};

export const reducers: any = combineReducers({
  formData,
  errorMessage
});

// store.js
export function configureStore(initState: any = {}): any {
  const storeInner: any = createStore(reducers, initState);
  return storeInner;
}

export const store: any = configureStore();
