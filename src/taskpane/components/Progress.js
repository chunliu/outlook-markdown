import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";
/* global Spinner */

export default class Progress extends React.Component {
  render() {
    const { logo, message, title } = this.props;

    return (
      <section className="ms-welcome__progress ms-u-fadeIn500">
        <img width="90" height="90" src={logo} alt={title} title={title} />
        <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
        <Spinner type={SpinnerSize.large} label={message} />
      </section>
    );
  }
}
