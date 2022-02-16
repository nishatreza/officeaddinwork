import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global Word, require */


const App = ({ title, isOfficeInitialized }) => {



  const click = async () => {
    return Word.run(async (context) => {

      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      {/* <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" /> */}
      {/* <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}> */}
      <p className="ms-font-l">
        Modify the source files, then click <b>Run</b>.
      </p>
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        Run
      </DefaultButton>
      {/* </HeroList> */}
    </div>
  );
}


App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;