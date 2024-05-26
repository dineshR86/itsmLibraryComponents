/* eslint-disable */
import * as React from "react";
import { Button, Checkbox,Field, Label, Select, Textarea,TextareaProps,makeStyles,shorthands,tokens, useId } from "@fluentui/react-components";
import {Save24Filled,ClearFormatting16Regular} from "@fluentui/react-icons";
import {SPHttpClient,SPHttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http';
import { ChatMessages } from "./ChatMessages";

interface TicketCloseProps{
    itemId:number;
    spHttpClient:SPHttpClient;
    listName:string;
    siteUrl: string;
    data?:any;
}

interface Iaddnotes {
    CloseRemark?: boolean
    ITSM360_UsersToNotifyId?: { results: number[] };
    Note?: string;
    NoteAuthorId?: number;
    RelatedItemId?: number;
    RelatedList?: string;
    TeamToNotifyId?: number;
    Title?: string;
    __metadata?: { type: string }


    CommunicationInitiatorId?: number;
    Communications?: string;
    FromMyIT?: boolean;
}



const useStyles = makeStyles({
    base: {
      display: "flex",
      flexDirection: "column",
      rowGap: tokens.spacingVerticalMNudge,
    },
    drpStatus:{
        display:"grid",
        gridTemplateRows: "repeat(1fr)",
        justifyItems: "start",
        ...shorthands.gap("2px"),
        maxWidth: "400px",
    },
    btnWrapper:{
        columnGap: "15px",
        display: "flex"
    }
  });

export const TicketClose:React.FC<TicketCloseProps>=(props)=>{
    const styles=useStyles();
    const {spHttpClient,itemId,siteUrl,listName,data}=props;
    const closeStatus=["","Awaiting Confirmation","Awaiting Requestor","Closed","Closed - Forwarded"];
    const dropdownId = useId("dropdown-default");

    // State Object Start
    const [closingComments,setClosingComments]=React.useState("");
    const [notifyRequestor,setNotifyRequestor]=React.useState(false);
    const [firstTimeResolution,setFirstTime]=React.useState(false);
    const [notifyAgent,setAgent]=React.useState(false);
    const [notifyTeam,setTeam]=React.useState(false);
    const [notifyStaff,setStaff]=React.useState(false);
    const [prevNotes,setPrevNotes]=React.useState<any[]>([]);
    const [assignedPersonId,setAssignedPersonID]=React.useState<number>(0);
    const [assignedTeam,setAssignedTeam]=React.useState<number>(0);
    const [assignedRequestor,setrequestor]=React.useState<number>(0);
    const [selectedcloseStatus,setCloseStatus]=React.useState("");
    const [currentUserId,setCurrentUserID]=React.useState<number>(0);

    // State Object End
    
    React.useEffect(()=>{
        getTicketNotes();
        getTicketCommunication();
        getCurrentUser();
        if (data && data.AssignedPerson && data.AssignedPerson.Id && data.AssignedPerson.Id !== 0) {
            setAssignedPersonID(data.AssignedPerson.Id);
        }

        if (data && data.AssignedTeam && data.AssignedTeam.Id && Number(data.AssignedTeam.Id) !== 0) {
            setAssignedTeam(data.AssignedTeam.Id);
        }

        if (data && data.Requester && data.Requester.Id && Number(data.Requester.Id) !== 0) {
            setrequestor(data.Requester.Id);
        }

        
    },[]);

    const getTicketNotes=async ()=>{
            const selectedFields = "$select=Note,NoteAuthor/EMail,NoteAuthor/Title,Created,TeamToNotify/Id,TeamToNotify/Title,ITSM360_UsersToNotify/Id,ITSM360_UsersToNotify/Title";
            const expandFields = "$expand=NoteAuthor,TeamToNotify,ITSM360_UsersToNotify";
            const filterQuery = "$filter=RelatedItemId eq '" + itemId + "' and RelatedList eq '" + listName + "' and CloseRemark eq 1";
            const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Notes')/items?${selectedFields}&${expandFields}&${filterQuery}&$top=5000`;
            const rawResponse:SPHttpClientResponse=await spHttpClient.get(url,SPHttpClient.configurations.v1);
            const x=await rawResponse.json();
            //console.log("notes data: ",x);
            let y:any[]=prevNotes;
            x.value.forEach((i:any)=>{
            const z:string=i.NoteAuthor.Title;
            y.push({
              UserDisplayName:z,
              Created:i.Created,
              Note:i.Note,
              UserInitials:`${z.split(" ")[0].substring(0,1)}${z.split(" ")[1].substring(0,1)}`
            });
          });
          

          //let z=[...prevNotes,...y];
          console.log("notes data: ",y);
          setPrevNotes(y);
    }

    const getTicketCommunication=async ()=>{
        const selectedFields = "$select=Communications,CommunicationInitiator/Title,CommunicationInitiator/EMail,Created";
        const expandFields = "$expand=CommunicationInitiator";
        const filterQuery = "$filter=RelatedItemId eq '" + itemId + "' and RelatedList eq '" + listName + "' and CloseRemark eq 1";
        const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket communications')/items?${selectedFields}&${expandFields}&${filterQuery}&$top=5000`;
        const rawResponse:SPHttpClientResponse=await spHttpClient.get(url,SPHttpClient.configurations.v1);
        const x=await rawResponse.json();
        
        let y:any[]=prevNotes;
        x.value.forEach((i:any)=>{
            const z:string=i.CommunicationInitiator.Title;
            y.push({
              UserDisplayName:z,
              Created:i.Created,
              Note:i.Communications,
              UserInitials:`${z.split(" ")[0].substring(0,1)}${z.split(" ")[1].substring(0,1)}`
            });
          });
        
          //let z=[...prevNotes,...y];
          console.log("communication data: ",y);
          setPrevNotes(y);
    }

    const getCurrentUser=async ()=>{
        const url = `${siteUrl}/_api/web/currentuser`;
        const rawResponse:SPHttpClientResponse=await spHttpClient.get(url,SPHttpClient.configurations.v1);
        const x=await rawResponse.json();
        setCurrentUserID(x.Id);
    }

    const submitClosingNotes=async ()=>{
        const newCommunicationData:Iaddnotes={};
        newCommunicationData.Title=data!=undefined?data.Title:"Close Notes";
        newCommunicationData.RelatedItemId=itemId;
        newCommunicationData.RelatedList=listName;
        newCommunicationData.Communications=closingComments;
        newCommunicationData.CloseRemark=true;
        newCommunicationData.CommunicationInitiatorId=currentUserId;
        newCommunicationData.__metadata = { "type": "SP.Data.TicketCommunicationsListItem" };
        const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Communications')/items`;
        submitPostRequest(url,newCommunicationData);

        const addNotes: Iaddnotes = {};
        addNotes.__metadata = { "type": "SP.Data.TicketNotesListItem" };
        addNotes.Title = data!=undefined?data.Title:"Close Notes";
        addNotes.Note = closingComments;
        addNotes.NoteAuthorId = currentUserId;
        addNotes.RelatedList = listName;
        addNotes.RelatedItemId = itemId;
        addNotes.CloseRemark = true;
        const notesurl = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Notes')/items`;
        submitPostRequest(notesurl,addNotes);
        setClosingComments("");

    }

    const submitPostRequest=async (url:string,data:any)=>{
        console.log("post url: ",url);
        const httpclientoptions: ISPHttpClientOptions = {
        headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": "",
            'X-HTTP-Method': 'POST'
        },
        body: JSON.stringify(data)
        };

        const rawResponse:SPHttpClientResponse= await spHttpClient.post(url,SPHttpClient.configurations.v1,httpclientoptions);
        console.log("Submit data: ",await rawResponse.json());
    }

    const onClosingNotesChange: TextareaProps["onChange"] = (ev, data) => {
        setClosingComments(data.value);
    };

    // const onActiveOptionChange = React.useCallback(
    //     (_, data) => {
    //         setCloseStatus(data?.nextOption?.text);
    //     },
    //     [setCloseStatus]
    //   );

    const onCloseStatusChange=(_ev:any)=>{
        console.log(_ev.target.value);
        setCloseStatus(_ev.target.value);
    };



    return(
        <div className={styles.base}>
          <Field size="large" label="Add short description for this tab (optional)Add short description for this tab (optional)Add short description for this tab (optional)">
            <Textarea onChange={onClosingNotesChange} placeholder="type here..." value={closingComments} />
          </Field>
          <br/>
          <div className={styles.drpStatus}>
                <Label id={dropdownId}>Select Status</Label>
                {/* <Dropdown onActiveOptionChange={onActiveOptionChange}>
                    {closeStatus.map((i)=>(
                        <option key={i}>{i}</option>
                    ))}
                </Dropdown> */}
                <Select id={dropdownId} value={selectedcloseStatus} onChange={onCloseStatusChange}>
                    {closeStatus.map((i)=>(
                        <option value={i}>{i}</option>
                    ))}
                </Select>
          </div>
          <br/>
          <div className={styles.base}>
                <Checkbox checked={firstTimeResolution} label="First Time Resolution" onChange={()=>setFirstTime((checked)=>!checked)} />
                <Checkbox disabled={assignedRequestor!=0?false:true} checked={notifyRequestor} label="Notify Requestor" onChange={()=>setNotifyRequestor((checked)=>!checked)} />
                <Checkbox disabled={assignedPersonId!=0?false:true} checked={notifyAgent} label="Notify Agent" onChange={()=>setAgent((checked)=>!checked)} />
                <Checkbox disabled={assignedTeam!=0?false:true} checked={notifyTeam} label="Notify Team" onChange={()=>setTeam((checked)=>!checked)} />
                <Checkbox checked={notifyStaff} label="Notify Staff Members" onChange={()=>setStaff((checked)=>!checked)} />
          </div>

          <div className={styles.btnWrapper}>
                <Button icon={<ClearFormatting16Regular />}>Clear</Button><Button size="large" icon={<Save24Filled />} onClick={submitClosingNotes}>Send and Inform</Button>        
          </div>

          <div>
            <ChatMessages messages={prevNotes} />
          </div>
          
        </div>
    )
}