import * as React from "react";
import { Provider, Flex, Text, Button, Header, Alert } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";


/**
 * Implementation of the emsignerTest Tab content page
 */
export const EmsignerTestTab = () => {


    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [Token, setToken] = useState('')


    useEffect(() => {
        console.log('test');
        
        if (inTeams === true) {
            microsoftTeams.appInitialization.notifySuccess();
        } else {
            console.log("Not in Microsoft Teams");
            if (context) { setEntityId(JSON.stringify(context)); }   // trial 1  delete later
        }
    }, [context]);   // in teams

    useEffect(() => {
        if (context) {
            setEntityId(JSON.stringify(context));
        }
    }, [context]);

    /**
     * The render() method to create the UI of the tab
     */

    const authenticate = () => {

        microsoftTeams.authentication.authenticate({

            url: window.location.origin + "/auth-start.html",
            // width: 600,
            // height: 535,
            successCallback: (accessToken: string) => {
                alert(`success--${accessToken}`);
                setToken(accessToken);
                // <Alert content={accessToken} dismissAction="Input alarm" header="TOKEN:" visible />
            },
            failureCallback: (reason) => {
                alert(`failed--${reason}`);
                // <Alert content={reason} dismissAction="Input alarm" header="Reason" visible />
            }
        });

    }


    const navigate = () => {
        if (Token) {
            alert("navigate")
        } else {
            alert('you cant login')
        }
    }


    return (

        <Provider theme={theme}>
            <Flex fill={true} column styles={{ padding: ".8rem 0 .8rem .5rem" }}>

                <Flex.Item>
                    <Header content="Welcome  to Emsigner" />
                </Flex.Item>


                <Flex.Item >
                    <div >
                        {!Token && <Button onClick={authenticate} styles={{ margin: "5rem 5rem" }}>sign in</Button>}
                        <Button onClick={navigate} styles={{ margin: "5rem 5rem" }}>Login to navigate</Button>

                        <br />
                        {Token && <Button>
                            <a href="https://qa-int.emsigner.com/eMsecure/Flexiform?DocID=d2XQz0fPXAHOzBcPypj6UQ%253d%253d">  Re-directional URL</a>
                        </Button>}

                    </div>
                </Flex.Item>


                <Flex.Item styles={{ padding: ".8rem 0 .8rem .5rem" }}>
                    <Text size="smaller" content="(C) Copyright emsigner" />
                </Flex.Item>

            </Flex>
        </Provider>
    );
};
