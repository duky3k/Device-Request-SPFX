/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/no-unescaped-entities */
import * as React from "react";
import { IFormRequestProps } from "./IFormRequestProps";
import { TextField } from "@fluentui/react/lib/TextField";
import { Stack, IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
import {
  Dropdown,
  // IDropdownStyles,
  IDropdownOption,
} from "@fluentui/react/lib/Dropdown";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  DatePicker,
  Dialog,
  DialogFooter,
  DialogType,
  PrimaryButton,
} from "office-ui-fabric-react";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/Peoplepicker";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { IFormRequestStates } from "./IFormRequestStates";
import styles from "./FormRequest.module.scss";
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 5 },
  styles: { root: { width: "50%" } },
};
// const dropdownStyles: Partial<IDropdownStyles> = {
//   dropdown: { width: 330 },
// };

const options: IDropdownOption[] = [
  { key: "1", text: "Monitor" },
  { key: "2", text: "PC" },
  { key: "3", text: "Laptop" },
  { key: "4", text: "Headphone" },
  { key: "5", text: "Screen" },
];
const optionsManufactuer: IDropdownOption[] = [
  { key: "1", text: "Dell" },
  { key: "2", text: "Asus" },
  { key: "3", text: "Hp" },
  { key: "4", text: "Acer" },
  { key: "5", text: "Apple" },
];
const optionsDu: IDropdownOption[] = [
  { key: "1", text: "DU3" },
  { key: "2", text: "DU4" },
  { key: "3", text: "DU11" },
  { key: "4", text: "DU21" },
];
const optionsSite: IDropdownOption[] = [
  { key: "1", text: "VN - HN 789 Hoàng Quốc Việt" },
  { key: "2", text: "HN CMC Tower" },
  { key: "3", text: "HN Việt Á" },
  { key: "4", text: "HN Bamboo" },
  { key: "5", text: "HN Peakview Tower" },
];

const dialogContent = {
  type: DialogType.normal,
  Title: "Message",
  subText: "Your request have been submitted successfully",
  closeButtonAriaLabel: "Close",
};
export default class FormRequest extends React.Component<
  IFormRequestProps,
  IFormRequestStates
> {
  private sp: SPFI;
  constructor(props: IFormRequestProps) {
    super(props);
    const char = "123456ABCDEFGHI762009765210BNMUOP";
    let randomNum = "";
    for (let index = 0; index < 17; index++) {
      const randomSignle = char[Math.floor(Math.random() * char.length)];
      randomNum += randomSignle;
    }

    this.state = {
      formData: {
        Name: "",
        Code: "",
        Serial: "DV" + randomNum,
        Site: "",
        From: "Wed Feb 22 2023",
        To: "Wed Feb 22 2023",
        Owner: this.props.userDisplayName,
        DU: "",
        Unit: "",
        Note: "",
        Group: "",
        Manufacturer: "",
        Supplier: "",
        purchaseDate: "",
        Price: "",
        Warranty: "",
        VAT: "",

        //hidenDialog:
      },
      hideDialog: true,
    };

    this.sp = spfi().using(SPFx({ pageContext: props.context.pageContext }));
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  public componentDidMount() {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ReadItem();
  }
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  public generatorNumber() {
    const char = "123456ABCDEFGHI76wndixdzsfszfs2009765210ojfngdBNMUOP";
    let randomNum = "";
    for (let index = 0; index < 20; index++) {
      const randomSignle = char[Math.floor(Math.random() * char.length)];
      randomNum += randomSignle;
    }
    console.log(randomNum);
  }
  public render(): React.ReactElement<IFormRequestProps> {
    const { formData } = this.state;
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    return (
      <div>
        <section>
          <h3>General</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField
                label="Code"
                id="code"
                value={formData.Code}
                required
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Code: value,
                    },
                  });
                }}
              />
              <div id="validation_Code" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
              <TextField
                label="Name"
                multiline
                autoAdjustHeight
                required
                value={formData.Name}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Name: value,
                    },
                  });
                }}
              />
              <div id="validation_Name" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
              <TextField
                label="Associated Asset "
                id="associatedAssset"
                underlined
                defaultValue="None"
              />
              <TextField
                label="Serial"
                id="serial"
                readOnly
                value={formData.Serial}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Serial: value,
                    },
                  });
                }}
              />
            </Stack>

            <Stack {...columnProps}>
              <TextField
                label="Type"
                id="type"
                disabled
                defaultValue="Physical Asset"
              />
              <Dropdown
                placeholder="Select an options"
                label="Group"
                options={options}
                // styles={dropdownStyles}
                required
                defaultValue={formData.Group}
                onChange={(e, option) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Group: option.text,
                    },
                  });
                }}
              />
              <div id="validation_Group" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
              <TextField
                label="Unit"
                id="unit"
                value={formData.Unit}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Unit: value,
                    },
                  });
                }}
              />
              <TextField
                label="Note"
                id="note"
                multiline
                autoAdjustHeight
                value={formData.Note}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Note: value,
                    },
                  });
                }}
              />
            </Stack>
          </Stack>
        </section>

        <section>
          <h3> Assignment</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField label="Status" placeholder="New" disabled />
              <TextField
                label="Owner"
                disabled
                value={this.props.userDisplayName}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Owner: value,
                    },
                  });
                }}
              />
              <Dropdown
                label="DU"
                placeholder="Select an options"
                required
                options={optionsDu}
                defaultValue={formData.DU}
                //styles={dropdownStyles}
                onChange={(e, option) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      DU: option.text,
                    },
                  });
                }}
              />
              <div id="validation_DU" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
            </Stack>

            <Stack {...columnProps}>
              <Dropdown
                label="Site"
                placeholder="Select an options"
                required
                options={optionsSite}
                defaultValue={formData.Site}
                //styles={dropdownStyles}
                onChange={(e, option) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Site: option.text,
                    },
                  });
                }}
              />
              <div id="validation_Site" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
              <DatePicker
                label="From"
                value={new Date(formData.From)}
                onSelectDate={(date) =>
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      From: date.toDateString(),
                    },
                  })
                }
              />
              <DatePicker
                label="To"
                placeholder="Select a date..."
                value={new Date(formData.To)}
                onSelectDate={(date) =>
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      To: date.toDateString(),
                    },
                  })
                }
              />
            </Stack>
          </Stack>
        </section>

        <section>
          <h3>Purchasing</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <Dropdown
                label="Manufacturer"
                placeholder="Select an options"
                required
                options={optionsManufactuer}
                defaultValue={formData.Manufacturer}
                //styles={dropdownStyles}
                onChange={(e, option) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Manufacturer: option.text,
                    },
                  });
                }}
              />
              <div
                id="validation_Manufacturer"
                className={styles.formValidation}
              >
                <span> You can't leave this blank</span>
              </div>
              <TextField
                label="Price(VND)"
                type="number"
                value={formData.Price}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Price: value,
                    },
                  });
                }}
              />
              <TextField
                label="Warranty (Months)"
                type="number"
                value={formData.Warranty}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Warranty: value,
                    },
                  });
                }}
              />
            </Stack>
            <Stack {...columnProps}>
              <PeoplePicker
                context={this.props.context}
                titleText={"Supplier"}
                placeholder={"Enter supplier name"}
                required
                onChange={(data) =>
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      Supplier: data.toString(),
                    },
                  })
                }
              />
              <div id="validation_Supplier" className={styles.formValidation}>
                <span> You can't leave this blank</span>
              </div>
              <DatePicker
                label="Purchase Date"
                placeholder="Select a date..."
                // value={new Date(formData.purchaseDate)}
                onSelectDate={(date) =>
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      purchaseDate: date.toDateString(),
                    },
                  })
                }
              />
              <TextField
                label="VAT RATE(%)"
                type="number"
                value={formData.VAT}
                onChange={(_, value) => {
                  this.setState({
                    ...this.state,
                    formData: {
                      ...this.state.formData,
                      VAT: value,
                    },
                  });
                }}
              />
            </Stack>
          </Stack>
        </section>
        <hr />
        <Dialog
          dialogContentProps={dialogContent}
          hidden={this.state.hideDialog}
        >
          <DialogFooter>
            <PrimaryButton text="Close" onClick={this.toggleDialog} />
          </DialogFooter>
        </Dialog>
        <div>
          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton
              {...stackTokens}
              text="Submit"
              allowDisabledFocus
              onClick={() => this.createItem()}
            />
            <PrimaryButton text="Cancel" onClick={() => this.cancelBtn()} />
          </Stack>
        </div>
      </div>
    );
  }

  //Create Item
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  private createItem = async () => {
    let validation: boolean = true;
    if (this.state.formData.Code === "") {
      validation = false;
      document
        .getElementById("validation_Code")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.Name === "") {
      validation = false;
      document
        .getElementById("validation_Name")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.Group === "") {
      validation = false;
      document
        .getElementById("validation_Group")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.Site === "") {
      validation = false;
      document
        .getElementById("validation_Site")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.Manufacturer === "") {
      validation = false;
      document
        .getElementById("validation_Manufacturer")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.DU === "") {
      validation = false;
      document
        .getElementById("validation_DU")
        .setAttribute("style", "display: block !important");
    }
    if (this.state.formData.Supplier === "") {
      validation = false;
      document
        .getElementById("validation_Supplier")
        .setAttribute("style", "display: block !important");
    }
    if (validation) {
      const { formData } = this.state;
      document
        .getElementById("validation_Code")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_Name")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_Group")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_Site")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_Manufacturer")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_DU")
        .setAttribute("style", "display: none");
      document
        .getElementById("validation_Supplier")
        .setAttribute("style", "display: none");
      try {
        const addItem = await this.sp.web.lists
          .getByTitle("Assets")
          .items.add({ ...formData })
          .then(() => {
            this.setState({
              hideDialog: false,
            });
          });
        console.log(addItem, formData);
      } catch (e) {
        console.error(e);
      }
    }
  };

  //Get Item by ID
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  private async _ReadItem() {
    //validation
    const item = await this.sp.web.lists
      .getByTitle("Assets")
      .items.getById(10)();
    console.log(item);
  }

  //toogle Dialog
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  public toggleDialog = () => {
    this.setState({
      hideDialog: true,
    });
  };
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  public cancelBtn = () => {
    this.setState({
      ...this.state,
      formData: {
        Name: "",
        Code: "",
        Serial: "",
        Site: "",
        From: "",
        To: "",
        Owner: "",
        DU: "",
        Unit: "",
        Note: "",
        Group: "",
        Manufacturer: "",
        Supplier: "",
        purchaseDate: "",
        Price: "",
        Warranty: "",
        VAT: "",
      },
    });
  };
}
