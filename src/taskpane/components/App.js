import * as React from "react";
import { PrimaryButton, DefaultButton, TextField, Label, Stack } from "office-ui-fabric-react";
import Showdown from "showdown";
import BorderWrapper from "react-border-wrapper";
import Progress from "./Progress";

const preStyle = "font-family: Consolas,Courier,monospace;\
  font-size: 85%;\
  background-color: #f6f8fa; \
  display: block; \
  line-height: 1.5; \
  overflow: auto; \
  padding: 10px; \
  border-radius: 6px; \
  word-break: normal; \
  border-style: solid; \
  border-width: 1px; \
  border-color: #a19f9d;";

const codeStyle = "font-family:Consolas,Courier,monospace;\
  font-size: 85%;\
  background-color: #f6f8fa;\
  border-radius: 3px;\
  padding: 2px 4px;\
  margin: 0;\
  border-style: solid; \
  border-width: 1px; \
  border-color: #a19f9d;";

const PreCodeExtension = function () {
  // Using this workaround because the app running in the Outlook client 
  // doesn't support regex lookbehind
  const precode = {
    type: 'output',
    filter: function(text) {
      // console.log(`extension text: ${text}`);
      text = text.replace(new RegExp(`((?!<pre>).{5}|^.{0,4})<code>`, 'g'), `$1<code style="${codeStyle}">`);
      text = text.replace(new RegExp(`<pre><code>`, 'g'), `<pre style="${preStyle}"><code>`);
      return text;
    }
  };

  return [precode];
}

const converter = new Showdown.Converter({ extensions: PreCodeExtension() });

export default class App extends React.Component {
  state = {markdown: '', htmlText: ''};

  clickInsert = async () => {
    var item = Office.context.mailbox.item;

    item.body.setSelectedDataAsync(
      this.state.htmlText,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed){
              console.log(asyncResult.error.message);
          }
          else {
          }
      });
  };

  clickClear = async () => {
    this.setState({ markdown: '', htmlText: '' });
  }

  onMarkdownChange = async (event, newValue) => {
    console.log("On Change: " + newValue);
    var html = converter.makeHtml(newValue);
    console.log("On Change: " + html);
    this.setState({markdown: newValue, htmlText: html});
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Loading..." />
      );
    }

    const stackTokens = { childrenGap: 20 };

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
            value={this.state.markdown}
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

        <Stack horizontal tokens={stackTokens} className="insertButton">
          <PrimaryButton text="Insert"
            iconProps={{ iconName: "ChevronRight" }}
            disabled={this.state.htmlText === ""}
            onClick={this.clickInsert} />
          <DefaultButton text="Clear" iconProps={{ iconName: "Clear" }} 
            disabled={this.state.htmlText === ""}
            onClick={this.clickClear} />
        </Stack>

      </div>
    );
  }
}
