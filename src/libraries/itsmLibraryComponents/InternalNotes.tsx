/* eslint-disable */
import * as React from "react";
import { Field, Textarea,TextareaProps,makeStyles,tokens,Checkbox, Button } from "@fluentui/react-components";
import {SPHttpClient, SPHttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http'
import { ChatMessages } from "./ChatMessages";

const useStyles = makeStyles({
    base: {
      display: "flex",
      flexDirection: "column",
      rowGap: tokens.spacingVerticalMNudge,
    },
  });

  export interface InternalNotesProps{
    spHttpClient:SPHttpClient;
    itemId:number;
    siteUrl:string;
  }
  
  export interface IinternalNotes {
    userId?: string;
    EditorEMail?: string;
    Editor?: string;
    Note?: string;
    Modified?: string;
    // TeamToNotify?: string;
    image?: string;
    documents?: [];
    personPictureUrl?: string;
    personName?: string;
    msgTime?: string;
    content?: string;
    AttachmentFiles?: []
    Attachments?: boolean;
    Cc?: string;
    Created?: string;
    Email?: string;
    ID?: number
    Id?: number;
    Message?: string;
    PlainTextMessage?: string;
    Read?: boolean;
    Received?: boolean;
    RelatedItem?: number;
    RelatedList?: string;
    Title?: string;
    attachements?: []
    communication?: string;
    date?: string | Date;
    itemID?: number;
    knowledgeArticles?: [];
    personid?: number;
    // eslint-disable-next-line @rushstack/no-new-null
    read?: string | null;
    siteUrl?: string;
    email?: string;
    itemId?: number;
    listUrl?: string;
    listItemId?: number
    type?: string;
    attachments?: [],
    id?: number;
    name?: string | number;
    NoteAuthorId?: number;
    RelatedItemId?: number;
    __metadata?: { type: string };
    TeamToNotifyId?: string | number;
    ITSM360_UsersToNotifyId?: { results: number[] }
  }

export const InternalNotes:React.FC<InternalNotesProps>=(props)=>{
    const {spHttpClient,itemId,siteUrl}=props;
    const styles=useStyles();
    const [notifyAgent,setAgent]=React.useState(false);
    const [notifyTeam,setTeam]=React.useState(false);
    const [notifyStaff,setStaff]=React.useState(false);
    const [IsassignedTeam,setIsAssignedTeam]=React.useState(true);
    const [note,setNote]=React.useState("");
    const [prevNotes,setPrevNotes]=React.useState<any[]>([]);

    React.useEffect(()=>{
      (async ()=>{
      const selectFields = "$select=Note,Author/Id,Author/Title,Author/EMail,Created,RelatedList,ITSM360_UsersToNotify/Id,ITSM360_UsersToNotify/Title,ITSM360_UsersToNotify/Mail,TeamToNotify/Id,TeamToNotify/Title";
      const expandFields = "$expand=Author,ITSM360_UsersToNotify,TeamToNotify";
      const filterQuery = "$filter=RelatedItemId eq '" + itemId + "'";
      const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Notes')/items?${selectFields}&${expandFields}&${filterQuery}&$top=5000`;
      const rawResponse:SPHttpClientResponse=await spHttpClient.get(url,SPHttpClient.configurations.v1);
      
      const x=await rawResponse.json();
      console.log("notes data: ",x);
      let y:any[]=[];
      x.value.forEach((i:any)=>{
        const z:string=i.Author.Title;
        y.push({
          UserDisplayName:z,
          Created:i.Created,
          Note:i.Note,
          UserInitials:`${z.split(" ")[0].substring(0,1)}${z.split(" ")[1].substring(0,1)}`
        });
      });
      console.log("notes data1: ",y);
      //y.sort((i,j)=>i.Created > j.Created)
      setPrevNotes(y);
      setIsAssignedTeam(false);
    })()
  },[]);

  const submitNotes=async ()=>{
    const newNoteData:IinternalNotes={};
    newNoteData.Title="Test Notes";
    newNoteData.RelatedItemId=itemId;
    newNoteData.RelatedList="ITSM360_Tickets";
    newNoteData.Note=note;
    newNoteData.__metadata = { "type": "SP.Data.TicketNotesListItem" };
    const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Notes')/items`;
    const httpclientoptions: ISPHttpClientOptions = {
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
        'X-HTTP-Method': 'POST'
      },
      body: JSON.stringify(newNoteData)
    };

    const rawResponse:SPHttpClientResponse= await spHttpClient.post(url,SPHttpClient.configurations.v1,httpclientoptions);
    console.log("Submit data: ",await rawResponse.json());
    setNote("");
  }

  const onChange: TextareaProps["onChange"] = (ev, data) => {
    setNote(data.value);
  };

  // const previousNotes=async ()=>{
  //   const selectFields = "$select=Note,Author/Id,Author/Title,Author/EMail,Created,RelatedList,ITSM360_UsersToNotify/Id,ITSM360_UsersToNotify/Title,ITSM360_UsersToNotify/Mail,TeamToNotify/Id,TeamToNotify/Title";
  //     const expandFields = "$expand=Author,ITSM360_UsersToNotify,TeamToNotify";
  //     const filterQuery = "$filter=RelatedItemId eq '" + itemId + "'";
  //     const url = `${siteUrl}/_api/web/lists/GetByTitle('Ticket Notes')/items?${selectFields}&${expandFields}&${filterQuery}&$top=5000`;
  //   const rawResponse:SPHttpClientResponse=await spHttpClient.get(url,SPHttpClient.configurations.v1);
  //   const x=await rawResponse.json();
  //   let y:any[]=[];
  //   y.push(x.value.map((i:any)=>{
  //     return {
  //       UserDisplayName:i.Author.Title,
  //       Created:i.Created,
  //       Note:i.Note
  //     }
  //   }));

  //   setPrevNotes(y);
  //   //console.log("list data: ",await rawResponse.json(),x);
  // }

    return (
        <div className={styles.base}>
          <Field size="large" label="Internal Notes" hint="Internal Notes for tickets">
            <Textarea onChange={onChange} placeholder="type here..." />
          </Field>

          <Checkbox checked={notifyAgent} label="Notify Agent" onChange={()=>setAgent((checked)=>!checked)} />
          <Checkbox disabled={!IsassignedTeam} checked={notifyTeam} label="Notify Team" onChange={()=>setTeam((checked)=>!checked)} />
          <Checkbox checked={notifyStaff} label="Notify Staff Members" onChange={()=>setStaff((checked)=>!checked)} />

          <div>
            <Button onClick={submitNotes}>Send</Button>
          </div>

          <div>
            <ChatMessages messages={prevNotes} />
          </div>
        </div>
    )

    
}

