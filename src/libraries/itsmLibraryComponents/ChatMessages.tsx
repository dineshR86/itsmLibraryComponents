/* eslint-disable */
import * as React from "react";
import { Persona } from "@fluentui/react-components";
const useStyles = {
    container : {
      //border: "2px solid #dedede",
      backgroundColor:"transparent",
      //borderRadius: "5px",
        padding:"10px",
        margin:"10px 0"
    },
    timeLeft:{
        Float:"left",
        color:"#999",
        marginLeft:"10px"
    },
    profileImg:{
        width:"32px",
        height:"32px",
        margin:"8px",
        borderRadius:"50%",
        border:"1px solid #fff",
        backgroundColor:"#ca5010",
        color:"#fff"
    }
  };

  export interface ChatMessageProps{
    UserDisplayName:string;
    UserImageUrl?:string;
    Created:string;
    Note:string;
    UserInitials:string;
  }

  export interface ChatMessages{
    messages:ChatMessageProps[];
  }

  export const ChatMessages:React.FC<ChatMessages>=(props)=>{

    return (
        <div>
            {props.messages.map((i)=>{
                return (
                        <div style={useStyles.container}>
                            {/* <div><span style={useStyles.profileImg}>{i.UserInitials}</span> <span>{i.UserDisplayName}</span>
                                <span style={useStyles.timeLeft}>{new Date(i.Created).toLocaleString()}</span>
                            </div> */}
                            <div>
                                <Persona name={i.UserDisplayName} secondaryText={new Date(i.Created).toLocaleString()} presence={{status:"available"}} />
                            </div>
                            <p><div dangerouslySetInnerHTML={{__html:i.Note}}></div></p>
                        </div>
                )
            })}
        </div>
    )
  }