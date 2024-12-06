import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
// import HeroList from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { insertText } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();

  React.useEffect(() => {
    const getAccessToken = async () => {
      try {
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

        console.log("MS Word Log", userTokenEncoded);
      } catch (err) {
        console.log("MS Word Error", err);
      }
    };

    getAccessToken();
  }, []);

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome Mukesh" />
      <TextInsertion insertText={insertText} />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
