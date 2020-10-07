import * as React from "react";
import { PrimaryButton, TextField } from "office-ui-fabric-react";
import Progress from "./Progress";
/* global Button, Header, HeroList, HeroListItem, Progress */

export default class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {mdText: ''};
  }

  click = async () => {
    /**
     * Insert your Outlook code here
     */
  };

  onMarkdownChange = async (event, newValue) => {
    this.setState({mdText: newValue});
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome__main">

          <TextField label="Input markdown here" multiline autoAdjustHeight
            onChange={this.onMarkdownChange} />

          <PrimaryButton
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Insert
          </PrimaryButton>

      </div>
    );
  }
}
