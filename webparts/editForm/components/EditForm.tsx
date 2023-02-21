import * as React from 'react';
import { IEditFormProps } from './IEditFormProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Stack, TextField, Label, IStackProps, IStackStyles, PrimaryButton } from 'office-ui-fabric-react';
import { IEditFormState } from './IEditFormState';
import { spfi, SPFI, SPFx } from '@pnp/sp';

const stackTokens = { childrenGap: 50 };
const iconProps = { iconName: 'Calendar' };
const stackStyles: Partial<IStackStyles> = { root: { width: '100%' } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 5 },
  styles: { root: { width: '50%' } },
};

export default class EditForm extends React.Component<IEditFormProps, IEditFormState > {
  private sp: SPFI;
  constructor(props: IEditFormProps) {
    super(props);
    this.state = {
      formData: {
        Name: "",
        Code: "",
        Serial: "",
        Site: "",
        From: "",
        To: "",
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
        Status: "",
        Owner: "",
        //hidenDialog:
      },
    };

    this.sp = spfi().using(SPFx({ pageContext: props.context.pageContext }));
  }
  componentDidMount(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._ReadItem();
  }
  //Get Item by ID
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  private async _ReadItem(){
    const urlParams = new URLSearchParams(window.location.search);
    const id = parseInt(urlParams.get("uid"));
    
    const item = await this.sp.web.lists
    .getByTitle("Assets")
    .items.getById(id)().then((res)=>{
      this.setState({...this.state, formData:{
        ...this.state.formData,
        Code: res.Code,
        Name: res.Name,
        Group: res.Group,
        Unit: res.Unit,
        Serial: res.Serial,
        Note: res.Note,
        Site: res.Site,
        From: res.From,
        To: res.To,
        DU: res.DU,
        Owner: res.Owner,
        Manufacturer: res.Manufacturer,
        Supplier: res.Supplier,
        Price: res.Price,
        purchaseDate: res.purchaseDate,
        Warranty: res.Warranty,
        VAT: res.VAT,
        Status: res.Status
        
      }})
    });
    console.log(item);
  }
  public render(): React.ReactElement<IEditFormProps> {
    
    return (
      <div>
        <section>
          <h3>General</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles} >
            <Stack {...columnProps}>
              <TextField 
                label="Code" 
                id="code"
                value={this.state.formData.Code}
                disabled
              />
              <TextField
                label="Name"
                multiline
                autoAdjustHeight
                required
                value={this.state.formData.Name}
                disabled
              />
              <TextField
                label="Associated Asset "
                id="associatedAssset"
                underlined
                defaultValue="None"
                disabled
              />
              <TextField 
                label="Serial" 
                id="serial"
                value={this.state.formData.Serial} 
                disabled
              />
            </Stack>

            <Stack {...columnProps}>
              <TextField 
                label="Type" 
                id="type" 
                disabled
                defaultValue='Physical Asset'
              />
              <TextField
                label="Group"
                id="Group"
                required
                value={this.state.formData.Group}
                disabled
              />
              <TextField
                label="Unit"
                id="unit"
                value={this.state.formData.Unit}
                disabled
              />
              {/* <Stack horizontal>
                <Toggle label="Hired Asset" defaultChecked />
                <Toggle label="Manage By Admin" defaultChecked  />
              </Stack> */}
              <TextField 
                label="Note" 
                id="note" 
                multiline 
                autoAdjustHeight
                value={this.state.formData.Note}
                disabled
              />
            </Stack>
          </Stack>
        </section>

        <section>
          <h3> Assignment</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField 
                label="Status"
                disabled 
                value={this.state.formData.Status}
              />
              <TextField 
                label="Owner" 
                disabled 
                value={this.state.formData.Owner}/>
              <TextField
                label="DU"
                required
                disabled
                value={this.state.formData.DU}
              />
            </Stack>

            <Stack {...columnProps}>
              <TextField
                label="Site"
                required
                disabled
                value={this.state.formData.Site}
              />
              <Stack>
                <Label> From</Label>
                <TextField
                  disabled
                  iconProps={iconProps}
                  value={this.state.formData.From}
                />
                <Label> To</Label>
                <TextField
                  disabled
                  iconProps={iconProps}  
                  value={this.state.formData.To}
                />
              </Stack>
            </Stack>
          </Stack>
        </section>

        <section>
          <h3>Purchasing</h3>
          <hr />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField
                label="Manufacturer"
                required
                disabled
                //styles={dropdownStyles}
                value={this.state.formData.Manufacturer}
                
              />
              <TextField 
                label="Price(VND)" 
                value={this.state.formData.Price}
                disabled
              />
                
              <TextField 
                label="Warranty (Months)" 
                value={this.state.formData.Warranty}
                disabled
              />
            </Stack>
            <Stack {...columnProps}>
              <TextField 
                label="Supplier"
                required
                disabled
                placeholder='Duc Nguyen'
              />
              <TextField 
                label="Purchase Date" 
                iconProps={iconProps}
                value={this.state.formData.purchaseDate}
                disabled
              />
              <TextField 
                label="VAT RATE(%)"
                value={this.state.formData.VAT}
                disabled
              />
            </Stack>
          </Stack>
        </section>
        <hr />
        {/* <Dialog 
          dialogContentProps={dialogContent}
          hidden={this.state.formData.hideDialog}
          onDismiss={this.toggleDialog}
        >
          <DialogFooter>
          <PrimaryButton 
            text="Close" 
            onClick={this.toggleDialog}/>
          </DialogFooter>
        </Dialog> */}
        <div>
          <Stack horizontal tokens={stackTokens}>
            <a href="https://5fzc7g.sharepoint.com/sites/DeviceRequestSystem/SitePages/All-Requests.aspx" >
              <PrimaryButton text="Back"/>
            </a>
          </Stack>
        </div>
      </div>
    );
  }
  

}
