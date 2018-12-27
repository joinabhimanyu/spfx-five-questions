import * as React from "react";
import "./InformaticaFiveQuestions.module.scss";
import { IInformaticaFiveQuestionsProps } from "./IInformaticaFiveQuestionsProps";
import {
  IInformaticaFiveQuestionsState,
  IInformaticaFiveQuestionsFormData,
  IInformaticaFiveQuestionsField
} from "./IInformaticaFiveQuestionsState";
import MockHttpClient from "./MockHttpClient";
import * as MSHttp from "@microsoft/sp-http";

import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import { Modal } from "office-ui-fabric-react/lib/Modal";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Label } from "office-ui-fabric-react/lib/Label";

import * as strings from "InformaticaFiveQuestionsWebPartStrings";
import { Dialog } from "@microsoft/sp-dialog";


export default class InformaticaFiveQuestions extends React.Component<
  any,
  IInformaticaFiveQuestionsState
  > {
  private _isMounted: boolean = false;
  private _componentRefs: Array<any> = [];
  private _saveButtonContainerRef: HTMLElement | null;

  constructor() {
    super();
    this.state = {
      validationMessage: "",
      showValidationMsg: false,
      validProps: true,
      modalVisible: false,
      readOnlyModal: false,
      ID: 0,
      title: "",
      imageFileName: "",
      imageUrl: "",
      authorName: ""
    };
  }

  /**
   * begin section lifecycle events
   */
  public componentDidMount(): void {
    this._isMounted = true;
    if (this.validateProps() && this._isMounted) {
      this.setState({ validProps: true });
      this.getFields();
      this._getDisplyItem();

    } else {
      if (this._isMounted) {
        this.setState({ validProps: false });
      }
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }
  /**
   * end section lifecycle events
   */

  private validateProps(): boolean {
    if (
      !this.props.WebPartTitle ||
      !this.props.context ||
      !this.props.ListName
    ) {
      return false;
    } else {
      return true;
    }
  }

  private getFields(): void {
    if (Environment.type === EnvironmentType.Local) {
      this._getMockFields().then(response => {
        if (this._isMounted) {
          this.props.setFormData({ fields: response, attachement: null });
        }
      });
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      this._getFields().then(response => {
        if (this._isMounted) {
          this.props.setFormData({ fields: response, attachement: null });
        }
      });
    }
  }

  private _getMockFields(): Promise<IInformaticaFiveQuestionsField[]> {
    return MockHttpClient.getMockFields().then(
      (data: IInformaticaFiveQuestionsField[]) => {
        return data;
      }
    ) as Promise<IInformaticaFiveQuestionsField[]>;
  }

  private _getFields(): Promise<IInformaticaFiveQuestionsField[]> {

    const url: string = `/_api/web/Lists('${this.props.ListName}')/fields?` +
      `$select=ID,InternalName,Title,Required,ReadOnlyField,ValidationFormula,ValidationMessage,Value` +
      `&$filter=Hidden eq false and ReadOnlyField eq false`;

    return this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl + url,
        MSHttp.SPHttpClient.configurations.v1
      )
      .then((response: MSHttp.SPHttpClientResponse) => {
        return response.json();
      })
      .then(data => {
        const filteredFields: any = data.value.filter(item => {
          return item.InternalName !== "ContentType" && item.InternalName !== "ShowOnHomepage" && item.InternalName !== "Attachments";
        });
        return filteredFields.map(item => {
          return {
            Id: item.Id,
            ReadOnlyField: JSON.parse(item.ReadOnlyField),
            Required: JSON.parse(item.Required),
            StaticName: item.InternalName,
            DisplayName: item.Title,
            ValidationFormula: item.ValidationFormula,
            ValidationMessage: item.ValidationMessage,
            Value: ""
          };
        });
      }).catch((err: any) => {
        this.props.setErrMsg(err);
      });
  }

  private _getDisplyItem(): void {
    const url: string = `/_api/web/Lists('${this.props.ListName}')/Items?` +
      `$filter=ShowOnHomepage  eq '1' and Attachments eq '1'&$OrderBy=Modified desc&expand=Author&$top=1`;

    this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl + url,
        MSHttp.SPHttpClient.configurations.v1
      )
      .then((response: MSHttp.SPHttpClientResponse) => {
        return response.json();
      })
      .then(data => {
        const item: any = data.value[0];
        this._getDisplayItemAuthor(item.ID);
        this._getDisplayItemAttachments(item.ID, item.Title);
      }).catch((err: any) => {
        this.props.setErrMsg(err);
      });
  }

  private _getDisplayItemAuthor(itemId: string): void {
    const url: string = `/_api/web/Lists('${this.props.ListName}')/Items('${itemId}')?$select=Author/Title,Author/Name&$expand=Author`;
    this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl + url,
        MSHttp.SPHttpClient.configurations.v1
      )
      .then((response: MSHttp.SPHttpClientResponse) => {
        return response.json();
      })
      .then(data => {
        if (this._isMounted) {
          this.setState(
            {
              authorName: data.Author.Title
            }
          );
        }
      }).catch((err: any) => {
        this.props.setErrMsg(err);
      });
  }

  private _getDisplayItemAttachments(itemId: string, itemTitle: string): void {
    const url: string = `/_api/web/Lists('${this.props.ListName}')/Items('${itemId}')/AttachmentFiles`;
    this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl + url,
        MSHttp.SPHttpClient.configurations.v1
      )
      .then((response: MSHttp.SPHttpClientResponse) => {
        return response.json();
      })
      .then(data => {
        if (this._isMounted) {
          this.setState(
            {
              ID: parseInt(itemId, 10),
              title: itemTitle,
              imageFileName: data.value[0].FileName,
              imageUrl: data.value[0].ServerRelativeUrl
            }
          );
        }
      }).catch((err: any) => {
        this.props.setErrMsg(err);
      });
  }

  private getFileBuffer(file: any): Promise<any> {
    return new Promise((resolve, reject) => {
      let reader: any = new FileReader();
      reader.onload = (e: any) => {
        resolve(e.target.result);
      };
      reader.onerror = (e: any) => {
        reject(e.target.error);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  private uploadattachment(id: number): Promise<string> {
    return new Promise((resolve, reject) => {
      if (Environment.type === EnvironmentType.Local) {
        return Promise.resolve("success");
      } else if (
        Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint
      ) {
        if (this.props.formData.attachement) {
          let file: any = this.props.formData.attachement;
          this.getFileBuffer(file).then((buffer: any) => {
            const spOpts: MSHttp.ISPHttpClientOptions = {
              body: buffer
            };
            const url: string =
              "/_api/web/lists('" +
              this.props.ListName +
              "')/items(" +
              id +
              ")/AttachmentFiles/add(FileName='" +
              file.name +
              "')";
            this.props.context.spHttpClient
              .post(
                this.props.context.pageContext.web.absoluteUrl + url,
                MSHttp.SPHttpClient.configurations.v1,
                spOpts
              )
              .then((response: MSHttp.SPHttpClientResponse) => {
                return response.json();
              })
              .then(data => resolve("success")).catch((err: any) => {
                this.props.setErrMsg(err);
              })
              .catch((err: any) => {
                this.props.setErrMsg(err);
              });
          });
        }
      }
    });
  }

  private _postFormData(data: any): Promise<number> {
    const url: string = `/_api/web/Lists('${this.props.ListName}')/Items`;
    const spOpts: MSHttp.ISPHttpClientOptions = {
      body: JSON.stringify(data)
    };

    if (Environment.type === EnvironmentType.Local) {
      return Promise.resolve(10);
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      return this.props.context.spHttpClient
        .post(
          this.props.context.pageContext.web.absoluteUrl + url,
          MSHttp.SPHttpClient.configurations.v1,
          spOpts
        )
        .then((response: MSHttp.SPHttpClientResponse) => {
          return response.json();
        })
        .then(resp => resp.Id || 0).catch((err: any) => {
          this.props.setErrMsg(err);
        });
    }
  }
  /**
   * begin section event handlers
   */
  private toggleModal = (): void => {
    if (this._isMounted) {
      this.setState({
        modalVisible: false,
        readOnlyModal: false,
        ID: 0,
        title: "",
        imageFileName: "",
        imageUrl: ""
      });
      this.props.resetFormData();
      this.getFields();
      this._getDisplyItem();
    }
  }

  private _showModal = () => {
    if (this._isMounted) {
      this.setState({ modalVisible: true, readOnlyModal: false });
    }
  }

  private _closeModal = (): void => {
    if (this._isMounted) {
      this.setState({
        modalVisible: false,
        readOnlyModal: false
      });
      this.props.resetFormData();
      this.getFields();
    }
  }


  private _getErrorMessage = (ValidationFormula: string, ValidationMessage: string, Required: boolean) => (value: any) => {
    let errMessage: string = "";
    this.setState({ showValidationMsg: false, validationMessage: "" });
    // let value: string = e.value;
    if (Required && value === "") {
      errMessage = strings.RequiredMessage;
    } else {
      // do validation with formula and message;
    }
    return errMessage;
  }
  private _changeFieldValue = (staticName: string) => (value: any) => {
    if (staticName) {
      let fields: IInformaticaFiveQuestionsField[] = this.props.formData.fields.map(
        field => {
          if (field.StaticName === staticName) {
            field.Value = value;
          }
          return field;
        }
      );
      if (this._isMounted) {
        this.props.setFormData({ fields: fields, attachement: this.props.formData.attachement });
      }
    }
  }
  private _changeFileSelection = (e: any) => {
    if (e.currentTarget && e.currentTarget.files && e.currentTarget.files.length > 0) {
      if (this._isMounted) {
        this.props.setFormData({ fields: this.props.formData.fields, attachement: e.currentTarget.files[0] });
      }
    }
  }
  private validateFields = (): [boolean, Array<any>] => {
    let tuple: any = [true, []];
    const filter: any = this.props.formData.fields.filter((item) => !item.Value);
    if (filter && filter.length > 0) {
      tuple[0] = false;
      tuple[1] = filter.map((item) => {
        return {
          fieldLabel: item.DisplayName,
          errMessage: strings.RequiredMessage
        };
      });
    }
    return tuple;
  }

  private validateAttachement = (): boolean => !this.props.formData.attachement ? false : true;

  private _saveModal = () => {
    let data: any = {
    };
    this.setState({ showValidationMsg: false, validationMessage: "" });
    if (this.props.formData.fields && this.props.formData.fields.length > 0) {
      const tuple: any = this.validateFields();
      if (tuple[0]) {
        if (this.validateAttachement()) {
          for (const field of this.props.formData.fields) {
            data[field.StaticName] = field.Value;
          }
          this._postFormData(data).then((id: number) => {
            this.uploadattachment(id).then(() => {
              this.toggleModal();
            }).catch((err: any) => {
              this.props.setErrMsg(err);
            });
          }).catch((err: any) => {
            this.props.setErrMsg(err);
          });
        } else {
          // dialog.alert(`Please select attachment`);
          this.setState({ showValidationMsg: true, validationMessage: strings.AttachmentValidation });
        }
      } else {
        // show toast message for all validation errors
        const errs: Array<any> = tuple[1];
        this.setState({ showValidationMsg: true, validationMessage: strings.FieldValidation });
      }
    }
  }
  private onImageClick = () => {
    const url: string = `/_api/web/Lists('${this.props.ListName}')/Items('${this.state.ID}')`;
    this.props.context.spHttpClient
      .get(
        this.props.context.pageContext.web.absoluteUrl + url,
        MSHttp.SPHttpClient.configurations.v1
      )
      .then((response: MSHttp.SPHttpClientResponse) => {
        return response.json();
      })
      .then(resp1 => {
        let mFields: any = this.props.formData.fields.map((field) => ({
          ...field,
          Value: resp1[field.StaticName]
        }));
        this.props.setFormData({ fields: mFields, attachement: null });
        if (this._isMounted) {
          this.setState({ modalVisible: true, readOnlyModal: true });
        }
      }).catch((err: any) => {
        this.props.setErrMsg(err);
      });
  }

  private getComponentRef = (component: any) => {
    const filter: any = this._componentRefs.filter((comp) => comp.props.label === component.props.label);
    if (filter.length === 0) {
      this._componentRefs.push(component);
    }
  }

  /**
   * end section event handlers
   */

  public render(): React.ReactElement<IInformaticaFiveQuestionsProps> {
    const { readOnlyModal, imageUrl, title, validProps, modalVisible, showValidationMsg, validationMessage, authorName } = this.state;

    const readOnlyModalHeaderText: string = `Submitted by ${authorName}`;

    let fileNameSpan: any = null;
    let errorMsg: any = null;
    if (readOnlyModal) {
      // image thumbnail
      fileNameSpan = <a target="_blank" href={imageUrl} style={{ textDecoration: "none" }}>
        <Image
          title={title}
          src={imageUrl}
          alt={""}
          height={200}
          imageFit={ImageFit.cover}
        />
      </a>;
    } else {
      fileNameSpan = <input type="file" id={`addAttachment`}
        disabled={readOnlyModal} onChange={this._changeFileSelection} />;
    }
    if (this.props.errorMessage) {
      errorMsg = <MessageBar
        messageBarType={MessageBarType.blocked}
        isMultiline={false}
        truncated={true}
        overflowButtonAriaLabel="Overflow"
      >
        {this.props.errorMessage.message.toString()}
      </MessageBar>;
    }
    if (errorMsg) {
      return (
        <div>
          {errorMsg}
        </div>
      );
    } else if (!validProps) {
      return (
        <MessageBar>{strings.MessageBoxLabel}</MessageBar>
      );
    } else {
      return (
        <div>
          <div className="ms-Grid">
            <a>
              <div className="ms-Grid-row image-container">
                <Image
                  title={title}
                  src={imageUrl}
                  alt={""}
                  height={240}
                  imageFit={ImageFit.cover}
                  onClick={this.onImageClick}
                />
                <span>{title}</span>
              </div>
            </a>
            <div className="ms-Grid-row">
              {/* <DefaultButton onClick={this._showModal} text="Open Modal" /> */}
              <div className="submit-yours-button-container">
                <a className="ms-fontWeight-semibold" onClick={this._showModal}>
                  Submit Your's >
            </a>
              </div>
            </div>
          </div>
          <Modal
            isOpen={modalVisible}
            onDismiss={this._closeModal}
            isBlocking={false}
            containerClassName="ms-modalExample-container"
          >
            <div className="ms-modalExample-header">
              {readOnlyModal ? <span>{readOnlyModalHeaderText}</span> : <span>Submit Your's</span>}
              {/* <span>Submit Your's</span> */}
              <i className="ms-Icon ms-Icon--ChromeClose x-hidden-focus" aria-hidden="true" onClick={this._closeModal}></i>
            </div>
            <div className="ms-modalExample-body">
              <div className="docs-TextFieldErrorExample">

                <div className="field-Container">
                  <div className="ms-TextField root_049edaa7">
                    <div className="ms-TextField-wrapper wrapper_049edaa7">
                      <label className="ms-Label root-106">{readOnlyModal ? "" : "Upload Image"}</label>
                      <div className="field-attachment-input">
                        {fileNameSpan}
                      </div>
                    </div>
                  </div>
                </div>

                {/* repeat over fields */}
                {this.props.formData.fields.map((field, index) => (
                  <div key={index}>
                    {readOnlyModal ? <div className="view-field-container">
                      <Label className="question">{field.DisplayName}</Label>
                      <Label className="answer">{field.Value}</Label>
                      <hr />
                    </div> :
                      <div className="field-Container">
                        <TextField
                          underlined
                          componentRef={this.getComponentRef}
                          label={field.DisplayName}
                          value={field.Value}
                          required={field.Required}
                          onChanged={this._changeFieldValue(field.StaticName)}
                          onGetErrorMessage={this._getErrorMessage(field.ValidationFormula, field.ValidationMessage, field.Required)}
                          deferredValidationTime={2000}
                          validateOnLoad={true}
                          validateOnFocusIn={true}
                          validateOnFocusOut={true}
                        />
                      </div>
                    }
                  </div>
                ))}

                <div className="submit-button-container ms-Grid-row">
                  <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
                    {showValidationMsg ?
                      <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}>
                        {validationMessage}
                      </MessageBar>
                      : null}
                  </div>
                  <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg2">
                    {readOnlyModal ? null :
                      <div className="ms-CalloutExample-buttonArea"
                        ref={menuButton => (this._saveButtonContainerRef = menuButton)}>
                        <DefaultButton
                          disabled={readOnlyModal}
                          primary={true}
                          data-automation-id="submitYours"
                          text="Submit"
                          onClick={this._saveModal}
                        />
                      </div>
                    }
                  </div>
                </div>
              </div>
            </div>
          </Modal>
        </div>
      );
    }
  }
}
