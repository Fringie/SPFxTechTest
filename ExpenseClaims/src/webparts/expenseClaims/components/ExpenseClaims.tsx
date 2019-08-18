import * as React from 'react';
import styles from './ExpenseClaims.module.scss';
import { IExpenseClaimsProps } from './IExpenseClaimsProps';
import {sp} from "@pnp/sp";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PrimaryButton} from "office-ui-fabric-react";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Link } from 'office-ui-fabric-react/lib/Link';

type IExpenseClaimsState = {
  firstName: string;
  lastName: string;
  expenseDate: Date;
  expenseDescription: string;
  expenseCost: number;
  firstDayOfWeek?: DayOfWeek;
  sentRequest: boolean;
  sendingRequest: boolean;
};


const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker',

  isRequiredErrorMessage: 'Expense date is required.',

  invalidInputErrorMessage: 'Invalid date format.'
};

export default class ExpenseClaims extends React.Component<IExpenseClaimsProps, IExpenseClaimsState> {
  
  constructor(props) {
    super(props);
    
    this.state = { 
        expenseDate: new Date(),
        firstName: this.props.context.pageContext.user.displayName.split(' ')[0],
        lastName: this.props.context.pageContext.user.displayName.split(' ').slice(-1)[0],
        expenseDescription: null,
        expenseCost: null,
        firstDayOfWeek: DayOfWeek.Monday,
        sentRequest: false,
        sendingRequest: false
    }; 
  }

  public render(): React.ReactElement<IExpenseClaimsProps> {
    return (
      <div className={ styles.expenseClaims }>
        {this.renderExpenseUploaded()}
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <form>
                {this.renderFormFields()}
                <PrimaryButton
                  onClick={this._onSubmitForm.bind(this)}
                  text="Submit"
                  disabled={!this.state.expenseCost || !this.state.expenseDate || !this.state.expenseDescription 
                    || !this.state.firstName || !this.state.lastName || this.state.sentRequest}
                />
                <Spinner label="Uploading expense..." 
                  className={this.state.sendingRequest && !this.state.sentRequest ? styles.displayFlex : styles.hide} 
                />
              </form>
              
            </div>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Render the expense pop up on successful upload
   */
  private renderExpenseUploaded(){
    return (
      <div>
        {
          this.state.sentRequest ?
          <MessageBar messageBarType={MessageBarType.success} >
            Successfully uploaded expense
            <Link href={location.href} className={styles.blueTxt}>
              Add another expense
            </Link>
          </MessageBar>
          :
          <div></div>
        }
      </div>
    );
  }

  /**
   * Render the fields that are contained within the form
   */
  private renderFormFields(){
    return (
      <div>
        <TextField label="First name:" 
          required 
          onChanged={this._onSetFirstName} 
          value={this.state.firstName}
          onGetErrorMessage={this._getErrorMessage}
          validateOnFocusOut
          validateOnLoad={false}
        />
        <TextField label="Last name:" 
          required
          onChanged={this._onSetLastName}
          value={this.state.lastName}
          onGetErrorMessage={this._getErrorMessage}
          validateOnFocusOut
          validateOnLoad={false}
        />
        <DatePicker
          label="Expense Date"
          isRequired
          placeholder="Select a date..."
          ariaLabel="Select a date"
          firstDayOfWeek={this.state.firstDayOfWeek}
          onSelectDate={this._onSelectExpenseDate}
          strings={DayPickerStrings}
          value={this.state.expenseDate}
        />
        <TextField 
          label="Expense Cost"
          prefix="£"
          ariaLabel="Expense Cost"
          type="number"
          required
          onChanged={this._onSetExpenseCost}
          onGetErrorMessage={this._getNumberErrorMessage.bind(this)}
          validateOnLoad={false}
          validateOnFocusOut
        />
        <TextField label="Expense description:"
          required 
          multiline
          autoAdjustHeight
          onChanged={this._onSetExpenseDescription}
          onGetErrorMessage={this._getErrorMessage}
          validateOnFocusOut
          validateOnLoad={false}
        />
      </div>
    );
  }

  /**
   * Upload form data to sharepoint
   */
   private _onSubmitForm(){
      this.setState({sendingRequest: true});    
      sp.web.lists.getByTitle(this.props.listName).items.add({
        Firstname: this.state.firstName,
        Lastname: this.state.lastName,
        Expense_x0020_Cost: this.state.expenseCost,
        Expense_x0020_Description: this.state.expenseDescription,
        Expense_x0020_Date: this.state.expenseDate
      }).then(res => {
       //console.log(res); // do nothing with response but would be nice to add a check to ensure data has been uploaded correctly
       this.setState({sentRequest: true});
       this.setState({sendingRequest: false});
      });
   }

   private _onSetFirstName = (value: string | null | undefined): void => { 
    this.setState({firstName: value});
   }

  private _onSetLastName = (value: string | null | undefined): void => { 
    this.setState({lastName: value});
  }

  private _onSelectExpenseDate = (date: Date | null | undefined): void => { 
    this.setState({ expenseDate: date }); 
  }

  private _onSetExpenseCost = (value: number | null | undefined): void => {
    this.setState({ expenseCost: value});
  }

  private _onSetExpenseDescription = (value: string | null | undefined): void => {
    this.setState({expenseDescription: value});
  }


  private _getErrorMessage = (value: string): string => {
    return value.length < 1 ? 'Please enter a value' : "";
  }

  private _getNumberErrorMessage = (value: number | null | undefined): string => {
    return (value === null || value === undefined || value < 0.01 ) ? 'Please enter a value of £0.01 or more' : "";
  }
}
