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


  const click = async () => {
    return Word.run(async (context) => {

      const insertOnlineVideo = (vdName, vdHtml, vdWidth, vdHeight) => {
        if (Office.CoercionType.Ooxml) {
          var xmlCode = replaceCode(vdName, escapeHtml(vdHtml), vdWidth, vdHeight);
          Office.context.document.setSelectedDataAsync(
            xmlCode,
            { coercionType: "ooxml" }
          );
        }
      }


      var vdName = "My Video";
      var vdHtml = '<iframe width="800" height="600" src="http://www.youtube.com/embed/qk51u8-4uo4" frameborder="0" allowfullscreen></iframe>'; //YouTube embed video code
      //var vdHtml = '<object width="800" height="600"><param name="movie" value="http://i.d.com.com/av/video/embed/player.swf" /><param name="background" value="#333333" /><param name="allowFullScreen" value="true" /><param name="allowScriptAccess" value="true" /><param name="FlashVars" value="playerType=embedded&type=id&value=50139900" /><embed src="http://www.cnet.com/av/video/embed/player.swf" type="application/x-shockwave-flash" background="#333333" width="800" height="600" allowFullScreen="true" allowScriptAccess="always" FlashVars="playerType=embedded&type=id&value=50139900" /></object>'; //CNET's embed video code
      var vdWidth = 800;
      var vdHeight = 600;
      insertOnlineVideo(vdName, vdHtml, vdWidth, vdHeight);

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