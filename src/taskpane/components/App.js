import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global Word, require */


const App = ({ title, isOfficeInitialized }) => {
  const listItems = [
    {
      icon: "Ribbon",
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: "Unlock",
      primaryText: "Unlock features and functionality",
    },
    {
      icon: "Design",
      primaryText: "Create and visualize like a pro",
    },
  ];

  const insertOnlineVideo = (vdName, vdHtml, vdWidth, vdHeight) => {
    if (Office.CoercionType.Ooxml) {
      var xmlCode = replaceCode(vdName, escapeHtml(vdHtml), vdWidth, vdHeight);
      Office.context.document.setSelectedDataAsync(
        xmlCode,
        { coercionType: "ooxml" }
      );
    }
  }

  const replaceCode = () => {
    var i = 0;
    var args = arguments;
    var str = $('#xmlcode').val();
    return str.replace(/@@/g, function () { return args[i++]; });
  }

  //Escape HTML
  const escapeHtml = (str) => {
    var elm = document.createElement('pre');
    if (typeof elm.textContent != 'undefined') {
      elm.textContent = str;
    } else {
      elm.innerText = str;
    }
    return elm.innerHTML.replace(/["']/g, '&quot;');
  }

  const click = async () => {
    return Word.run(async (context) => {



      // insert video
      // var vdName = "My Video";
      // var vdHtml = '<iframe width="800" height="600" src="http://www.youtube.com/embed/qk51u8-4uo4" frameborder="0" allowfullscreen></iframe>'; //YouTube embed video code
      // var vdWidth = 800;
      // var vdHeight = 600;
      // insertOnlineVideo(vdName, vdHtml, vdWidth, vdHeight);

      // insert image
      var body = context.document.body;

      //for online img
      // body.insertHtml(
      //   '<img src="https://i.insider.com/60ae5f85bee0fc0019d598e0?width=700" alt="Girl in a jacket" width="100" height="100">', Word.InsertLocation.start);

      // from src image
      body.insertHtml(
        `<img src=${require("./../../../assets/logo-filled.png")} alt="Girl in a jacket" width="100" height="100">`, Word.InsertLocation.start);

      // inert text

      // Create a proxy object for the document body.
      // var body = context.document.body;
      // body.insertHtml(
      //   '<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

      // insert paragraph
      // let secondSentence = context.document.body.insertParagraph(
      //   "This is the first text with a custom style.",
      //   "End"
      // );
      // secondSentence.font.set({
      //   bold: false,
      //   italic: true,
      //   name: "Berlin Sans FB",
      //   color: "blue",
      //   size: 30
      // });

      // insert table

      // Use a two-dimensional array to hold the initial table values.
      // let data = [
      //   ["Tokyo", "Beijing", "Seattle"],
      //   ["Apple", "Orange", "Pineapple"]
      // ];
      // let table = context.document.body.insertTable(2, 3, "Start", data);
      // table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
      // table.styleFirstColumn = false;

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
      <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems} />
      <p className="ms-font-l">
        Modify the source files, then click <b>Run</b>.
      </p>


      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
        Run
      </DefaultButton>

    </div>
  );
}


App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;