import * as React from "react";
import { PrimaryButton, TextField, Label, Stack } from "office-ui-fabric-react";
import Showdown from "showdown";
import BorderWrapper from "react-border-wrapper";
import Progress from "./Progress";

const converter = new Showdown.Converter();

export default class App extends React.Component {
  state = {htmlText: ''};

  click = async () => {
    var item = Office.context.mailbox.item;
    item.body.setSelectedDataAsync(
      this.state.htmlText,
      {
        coercionType: Office.CoercionType.Html
      },
      function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed){
              console.log(asyncResult.error.message);
          }
          else {
          }
      });
  };

  onMarkdownChange = async (event, newValue) => {
    console.log("On Change: " + newValue);
    var html = converter.makeHtml(newValue);
    console.log("On Change: " + html);
    this.setState({htmlText: html});
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Loading..." />
      );
    }

    return (
      <div className="ms-welcome__main">

        <BorderWrapper
          borderColour="#00bcf1"
          borderWidth="1px"
          borderRadius="15px"
          borderType="solid"
          innerPadding="4px"
          topElement={<Label>Markdown</Label>}
          topPosition={0.1}
          topOffset="15px"
          topGap="4px"
        >
          <TextField multiline autoAdjustHeight borderless className="markdowntf" resizable={false}
            onChange={this.onMarkdownChange} />
        </BorderWrapper>

        <BorderWrapper
          borderColour="#00bcf1"
          borderWidth="1px"
          borderRadius="15px"
          borderType="solid"
          innerPadding="4px"
          topElement={<Label>Preview</Label>}
          topPosition={0.1}
          topOffset="15px"
          topGap="4px"
        >
          <div className="preview-panel"
            dangerouslySetInnerHTML={{__html: this.state.htmlText}}></div>
        </BorderWrapper>

        <PrimaryButton className="insertButton"
          iconProps={{ iconName: "ChevronRight" }}
          disabled={this.state.htmlText === ""}
          onClick={this.click}
        >Insert</PrimaryButton>

      </div>
    );
  }
}
