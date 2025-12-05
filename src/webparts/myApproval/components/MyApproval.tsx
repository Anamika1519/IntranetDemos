import { escape, set } from "@microsoft/sp-lodash-subset";

import React, { useRef, useState } from "react";

import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";

import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";

import "bootstrap/dist/css/bootstrap.min.css";

import "../../../CustomCss/mainCustom.scss";

import "../components/MyApproval.scss";

import "bootstrap/dist/js/bootstrap.bundle.min.js";

import "../../../CustomJSComponents/CustomTable/CustomTable.scss";
// import "./CustomTable.scss";

import "../../verticalSideBar/components/VerticalSidebar.scss";
//import "./CustomTable.scss";
import { IMyApprovalProps } from "./IMyApprovalProps";

import Provider from "../../../GlobalContext/provider";

import UserContext from "../../../GlobalContext/context";

import CustomBreadcrumb from "../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb";
import { getType } from "../../../APISearvice/CustomService";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";

import {
  faEdit,
  faPaperclip,
  faSort,
  faEye,
  faTrashAlt,
} from "@fortawesome/free-solid-svg-icons";

import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";

import * as XLSX from "xlsx";
let currentItemID = "";
import moment from "moment";

// import {

//   addItem,

//   GetCategory,

//   getChoiceFieldOption,

//   getDiscussionComments,

//   getDiscussionFilter,

//   getDiscussionFilterAll,

//   getDiscussionForum,

//   getDiscussionMe,

//   getDiscussionMeAll,

//   updateItem,

// } from "../../../APISearvice/DiscussionForumService";

import { encryptId } from "../../../APISearvice/CryptoService";

import { getNews } from "../../../APISearvice/NewsService";

import Swal from "sweetalert2";

import { getCategory, getEntity } from "../../../APISearvice/CustomService";

import ReactQuill from "react-quill";

import { uploadFileToLibrary } from "../../../APISearvice/MediaService";

import "react-quill/dist/quill.snow.css";

import { SPFI } from "@pnp/sp/presets/all";

// import { fetchUserInformationList } from "../../../APISearvice/GroupTeamService";

// import Multiselect from "multiselect-react-dropdown";

import { getSP } from "../loc/pnpjsConfig";

import { Eye, Edit } from "react-feather";

import {
  getDataByID,
  getMyApproval,
  getMyRequest,
  updateItemApproval,
} from "../../../APISearvice/ApprovalService";
import DMSMyApprovalAction from "./DMSApprovalAction";
import { getApprovalListsData } from "../../../APISearvice/BusinessAppsService";
// import DMSMyFolderApprovalAction from "./DMSFolderApprovalAction";
import DMSMyFolderApprovalAction from "./DMSFolderApprovalAction1";
import Select from 'react-select';

interface ApprovalHierarchyItem {
  id?: number;
  level: string;
  approverRole: string;
  approver: string;
  approvers: string[]; // New array for multiple approvers
  approvalCriteria: string;
  serialNumber?: number;
  assignedTo?: any[];
}
interface UserOption {
  value: string;
  label: string;
  email: string;
}

let actingforuseremail: any
const MyApprovalContext = ({ props }: any) => {
  const sp: SPFI = getSP();
  const [activeComponent, setActiveComponent] = useState<string>("");
  const { useHide }: any = React.useContext(UserContext);
  const [showNestedDMSTable, setShowNestedDMSTable] = useState(false);
  const [announcementData, setAnnouncementData] = React.useState([]);

  const [myApprovalsData, setMyApprovalsData] = React.useState([]);
  const [myApprovalsDataAll, setMyApprovalsDataAll] = React.useState([]);
  const [myApprovalsDataAutomation, setMyApprovalsDataAutomation] =
    React.useState([]);
  const handleShowNestedDMSTable = () => {
    setShowNestedDMSTable(true); // Show nested table within DMS
  };
  const elementRef = React.useRef<HTMLDivElement>(null);

  const SiteUrl = props.siteUrl;
  const [folderActionOrFileAction, setFolderActionOrFileAction] = useState("");
  const [newsData, setNewsData] = React.useState([]);

  const [TypeData, setTypeData] = React.useState([]);

  const [BnnerImagepostArr, setBannerImagepostArr] = React.useState([]);

  const [DocumentpostArr, setDocumentpostArr] = React.useState([]);

  const [DocumentpostArr1, setDocumentpostArr1] = React.useState([]);

  const [ImagepostArr, setImagepostArr] = React.useState([]);

  const [ImagepostArr1, setImagepostArr1] = React.useState([]);

  const [GrouTypeData, setGroupTypeData] = React.useState([]);

  const [DocumentpostIdsArr, setDocumentpostIdsArr] = React.useState([]);

  const [selectedValue, setSelectedValue] = useState([]);

  const [EnityData, setEnityData] = React.useState([]);

  const [options, setOpions] = useState([]);

  const [approved, setApproved] = useState("yes");

  const [filters, setFilters] = React.useState({
    SNo: "",

    RequestID: "",

    ProcessName: "",

    RequestedBy: "",

    RequestedDate: "",

    Status: "",
    Title: "",
  });

  const [StatusTypeData, setStatusTypeData] = useState([
    { id: "Pending", name: "Pending" },
    { id: "Approved", name: "Approved" },
    { id: "Rejected", name: "Rejected" },
  ]);
  const [isOpen, setIsOpen] = React.useState(false);

  const [IsinvideHide, setIsinvideHide] = React.useState(false);
  const [Mylistdata, setMylistdata] = useState([]);
  const [ProjectWorkflowdata, setProjectWorkflowdata] = useState([]);
  const [selectedProjectTask, setSelectedProjectTask] = React.useState<any>(null);
  const [showProjectForm, setShowProjectForm] = React.useState(false);
  const [projectNeedsFurtherApproval, setProjectNeedsFurtherApproval] = React.useState<string>("Select");
  const [projectWantsToPublishInDossier, setProjectWantsToPublishInDossier] = React.useState<string>("Select");
  const [projectRemarks, setProjectRemarks] = React.useState<string>("");
  const [projectHierarchy, setProjectHierarchy] = React.useState<ApprovalHierarchyItem[]>([]);
  const [projectDocumentInfo, setProjectDocumentInfo] = React.useState<{
    documentUrl: string;
    fileName: string;
    fileLeafRef: string;
    fileRef: string;
    sharedLink?: string;
  } | null>(null);
  const [users, setUsers] = React.useState<UserOption[]>([]);

  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated");
  };
  // const getApprovalmasterTasklist = async () => {
  //   try {
  //     const items = await sp.web.lists
  //       .getByTitle("DMSFileApprovalTaskList")
  //       .items.select(
  //         "Log",
  //         "CurrentUser",
  //         "Remark",
  //         "LogHistory",
  //         "FileUID/FileUID",
  //         "FileUID/SiteName",
  //         "FileUID/DocumentLibraryName",
  //         "FileUID/FileName",
  //         "FileUID/RequestNo",
  //         // , "FileUID/FilePreviewUrl"
  //         "FileUID/Status",
  //         "FileUID/FolderPath",
  //         "FileUID/RequestedBy",
  //         "FileUID/Created",
  //         "FileUID/ApproveAction",
  //         "MasterApproval/ApprovalType",
  //         "MasterApproval/Level",
  //         "MasterApproval/DocumentLibraryName"
  //       )
  //       .expand("FileUID", "MasterApproval")
  //       .filter(`CurrentUser eq '${currentUserEmailRef.current}'`)();
  //     console.log(items, "DMSFileApprovalTaskList");
  //     setMylistdata(items);
  //   } catch (error) {
  //     console.error("Error fetching list items:", error);
  //   }
  // };
  // const MyDMSAPPROVALDATASTATUS = async (sp: any, value: any) => {
  //   try {
  //     const items = await sp.web.lists
  //       .getByTitle("DMSFileApprovalTaskList")
  //       .items.select(
  //         "Log",
  //         "CurrentUser",
  //         "Remark",
  //         "LogHistory",
  //         "FileUID/FileUID",
  //         "FileUID/SiteName",
  //         "FileUID/DocumentLibraryName",
  //         "FileUID/FileName",
  //         "FileUID/RequestNo",
  //         // , "FileUID/FilePreviewUrl"
  //         "FileUID/Status",
  //         "FileUID/FolderPath",
  //         "FileUID/RequestedBy",
  //         "FileUID/Created",
  //         "FileUID/ApproveAction",
  //         "MasterApproval/ApprovalType",
  //         "MasterApproval/Level",
  //         "MasterApproval/DocumentLibraryName"
  //       )
  //       .expand("FileUID", "MasterApproval")
  //       .filter(`CurrentUser eq '${currentUserEmailRef.current}' and FileUID/Status eq '${value}'`)();
  //     console.log(items, "DMSFileApprovalTaskList");
  //     setMylistdata(items);
  //   } catch (error) {
  //     console.error("Error fetching list items:", error);
  //   }
  // };
  const getUserTitleByEmail = async (userEmail: any) => {
    try {
      const user = await sp.web.siteUsers.getByEmail(userEmail)();
      return user.Title;
    } catch (error) {
      console.error("Error fetching user title:", error);
      return null;
    }
  };
  const [loading, setLoading] = useState(true);
  const [actingForUser, setSetActingForUser] = useState([]);

  const myActingfordata = async () => {
    try {
      const currentUserEmail = currentUserEmailRef.current;
      console.log("currentUserEmail myActingfordata", currentUserEmail);
      const today = new Date().toISOString();
      const delegateListItems = await sp.web.lists.getByTitle('ARGDelegateList').items.select(
        "DelegateName/EMail",
        "ActingFor/EMail",
        "ActingFor/Title",
        "DelegateName/Title",
        "StartDate",
        "EndDate",
        "Status"
      )
        .expand("DelegateName", "ActingFor")
        .filter(`ActingFor/EMail eq '${currentUserEmail}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)();

      console.log("delegateListItems myActingfordata", delegateListItems);

      // Extract unique ActingFor.Title and EMail values
      const uniqueTitlesAndEmails = [
        ...new Map(
          delegateListItems.map((item) => [item.DelegateName?.Title, { title: item.DelegateName?.Title, email: item.DelegateName?.EMail }])
        ).values(),
      ];

      // Set state with unique titles and emails
      setSetActingForUser(uniqueTitlesAndEmails.map((item, index) => ({ id: index.toString(), name: item.title, email: item.email })));
      console.log("setSetActingForUser", actingForUser);
      console.log("uniqueTitlesAndEmails", uniqueTitlesAndEmails);

    } catch (error) {
      console.error("Error fetching data:", error);
    } finally {
      //  setLoading(false);
    }
  };


  // const getApprovalmasterTasklist = async (value: any, actingfor: any) => {
  //   // alert(`Status value is ${value} is acting for ${actingfor} in DMS`)

  //   try {
  //     // Retrieve current user email
  //     const currentUserEmail = currentUserEmailRef.current;

  //     // Fetch the ARGDelegateList items where the current user is in the ActingFor column
  //     const today = new Date().toISOString(); // Get today's date in YYYY-MM-DD format
  //     // console.log("today", today);

  //     // const delegateListItems = await sp.web.lists.getByTitle('ARGDelegateList').items.select(
  //     //   "DelegateName/EMail",
  //     //   "ActingFor/EMail",
  //     //   "ActingFor/Title",
  //     //   "StartDate",
  //     //   "EndDate",
  //     //   "Status"
  //     // )
  //     // .expand("DelegateName", "ActingFor")
  //     // .filter(`ActingFor/EMail eq '${currentUserEmail}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)();

  //     // console.log("delegateListItems", delegateListItems);

  //     // const additionalFilters = delegateListItems.map((item:any) => `CurrentUser eq '${item.DelegateName.EMail}'`).join(' or ');
  //     // const combinedFilters = `CurrentUser eq '${currentUserEmail}'${additionalFilters ? ` or (${additionalFilters})` : ''} and FileUID/Status eq '${value}'`;

  //     // // Fetch items from DMSFileApprovalTaskList based on the combined filters
  //     // const items2 = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
  //     //   "Log", "CurrentUser", "Remark", "LogHistory", "FileUID/FileUID",
  //     //   "FileUID/SiteName", "FileUID/DocumentLibraryName", "FileUID/FileName",
  //     //   "FileUID/RequestNo", "FileUID/Processname", "FileUID/Status",
  //     //   "FileUID/FolderPath", "FileUID/RequestedBy", "FileUID/Created",
  //     //   "FileUID/ApproveAction", "MasterApproval/ApprovalType", "MasterApproval/Level",
  //     //   "MasterApproval/DocumentLibraryName"
  //     // )
  //     // .expand("FileUID", "MasterApproval")
  //     // .filter(combinedFilters)
  //     // .orderBy("Created", false)
  //     // .getAll();

  //     // console.log(items2, "DMSFileApprovalTaskList");


  //     let arr = [];
  //     let approvalData: any[] = [];
  //     if (!actingfor) {
  //       const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
  //         "Log", "CurrentUser", "Remark"
  //         , "LogHistory"
  //         , "FileUID/FileUID"
  //         , "FileUID/SiteName"
  //         , "FileUID/DocumentLibraryName"
  //         , "FileUID/FileName"
  //         , "FileUID/RequestNo"
  //         , "FileUID/Processname"
  //         //  ,"FileUID/FilePreviewUrl" 
  //         , "FileUID/Status"
  //         , "FileUID/FolderPath"
  //         , "FileUID/RequestedBy"
  //         , "FileUID/Created"
  //         , "FileUID/ApproveAction"
  //         , "MasterApproval/ApprovalType"
  //         , "MasterApproval/Level"
  //         , "MasterApproval/DocumentLibraryName"

  //       )
  //         .expand("FileUID", "MasterApproval")
  //         .filter(`CurrentUser eq '${currentUserEmailRef.current}' and FileUID/Status eq '${value}'`).orderBy("Created", false).getAll();
  //       console.log(items, "DMSFileApprovalTaskList");
  //       items.map((item) => {
  //         if (item.CurrentUser !== currentUserEmailRef.current) {
  //           arr.push(item)
  //           // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
  //         }

  //       });
  //       const updatedItems = await Promise.all(items.map(async (item) => {
  //         const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //         return { ...item, RequestedByTitle: requestedbyuserTitle };
  //       }));
  //       approvalData = updatedItems
  //       setMylistdata(updatedItems);
  //       return arr = updatedItems
  //     }
  //     if (actingfor !== "" && actingfor !== undefined) {
  //       const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
  //         "Log", "CurrentUser", "Remark"
  //         , "LogHistory"
  //         , "FileUID/FileUID"
  //         , "FileUID/SiteName"
  //         , "FileUID/DocumentLibraryName"
  //         , "FileUID/FileName"
  //         , "FileUID/RequestNo"
  //         , "FileUID/Processname"
  //         //  ,"FileUID/FilePreviewUrl" 
  //         , "FileUID/Status"
  //         , "FileUID/FolderPath"
  //         , "FileUID/RequestedBy"
  //         , "FileUID/Created"
  //         , "FileUID/ApproveAction"
  //         , "MasterApproval/ApprovalType"
  //         , "MasterApproval/Level"
  //         , "MasterApproval/DocumentLibraryName"

  //       )
  //         .expand("FileUID", "MasterApproval")
  //         .filter(`CurrentUser eq '${actingfor}' and FileUID/Status eq '${value}'`).orderBy("Created", false).getAll();
  //       console.log(items, "DMSFileApprovalTaskList");
  //       items.map((item) => {
  //         if (item.CurrentUser !== currentUserEmailRef.current) {
  //           arr.push(item)
  //           // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
  //         }

  //       });
  //       const updatedItems = await Promise.all(items.map(async (item) => {
  //         const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //         return { ...item, RequestedByTitle: requestedbyuserTitle };
  //       }));
  //       approvalData = updatedItems;
  //       setMylistdata(updatedItems);
  //       return arr = updatedItems
  //     }

  //     // const updatedItems2 = await Promise.all(items2.map(async (item) => {
  //     //   const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //     //   return { ...item, RequestedByTitle: requestedbyuserTitle };
  //     // }));
  //     if (!actingfor) {
  //       const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
  //         "*",
  //         "Folderdetail"
  //         , "Folderdetail/SiteTitle"
  //         , "Folderdetail/DocumentLibraryName"
  //         , "Folderdetail/CurrentUser"
  //         , "Folderdetail/FolderPath"
  //         , "Folderdetail/FolderName"
  //         , "Folderdetail/ParentFolderId"
  //         , "Folderdetail/Department"
  //         , "Folderdetail/Devision"
  //         , "Folderdetail/RequestNo"
  //         , "FolderMeta"
  //         , "FolderMeta/SiteName"
  //         , "FolderMeta/DocumentLibraryName"
  //         , "FolderMeta/ColumnName",
  //         "Folderdetail/Processname",
  //         "Folderdetail/Status",
  //         "Approver"
  //       ).expand("Folderdetail", "FolderMeta")
  //         .filter(`Approver eq '${currentUserEmailRef.current}' and Folderdetail/Status eq '${value}'`)();
  //       console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
  //       const normalizeItem2 = async (item: any) => ({
  //         Log: item?.Log || '', // Replace with appropriate mappings
  //         CurrentUser: item?.Folderdetail?.CurrentUser || '',
  //         Remark: item?.Remark || '',
  //         LogHistory: item?.LogHistory || '',
  //         // ProcessName:  item?.Folderdetail?.Processname,
  //         RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //         FileUID: {
  //           FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
  //           SiteName: item?.FolderMeta?.SiteName || '',
  //           DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
  //           // FileName: item?.FolderMeta?.FolderName || '',
  //           FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
  //           RequestNo: item?.Folderdetail?.RequestNo || '',
  //           Status: item?.Folderdetail?.Status || '',
  //           FolderPath: item?.Folderdetail?.FolderPath || '',
  //           // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
  //           // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //           Created: item?.Created || '',
  //           ApproveAction: item?.ApproveAction || '',
  //           Processname: item?.Folderdetail?.Processname
  //         },
  //         MasterApproval: {
  //           ApprovalType: item?.ApprovalType || '',
  //           Level: item?.Level || '',
  //           DocumentLibraryName: item?.DocumentLibraryName || ''
  //         }
  //       });
  //       // const normalizeItem3 = Item2.map(normalizeItem2);
  //       const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
  //       console.log("normalizeItem3", normalizeItem3);
  //       console.log(approvalData, "approvalData 666");
  //       const CombinedItems = [...approvalData, ...normalizeItem3];
  //       console.log(CombinedItems, "CombinedItems")
  //       setMylistdata(CombinedItems);
  //       // setMylistdata(updatedItems);


  //       return arr = CombinedItems
  //     }
  //     if (actingfor !== "" && actingfor !== undefined) {
  //       const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
  //         "*",
  //         "Folderdetail"
  //         , "Folderdetail/SiteTitle"
  //         , "Folderdetail/DocumentLibraryName"
  //         , "Folderdetail/CurrentUser"
  //         , "Folderdetail/FolderPath"
  //         , "Folderdetail/FolderName"
  //         , "Folderdetail/ParentFolderId"
  //         , "Folderdetail/Department"
  //         , "Folderdetail/Devision"
  //         , "Folderdetail/RequestNo"
  //         , "FolderMeta"
  //         , "FolderMeta/SiteName"
  //         , "FolderMeta/DocumentLibraryName"
  //         , "FolderMeta/ColumnName",
  //         "Folderdetail/Processname",
  //         "Folderdetail/Status",
  //         "Approver"
  //       ).expand("Folderdetail", "FolderMeta")
  //         .filter(`Approver eq '${actingfor}' and Folderdetail/Status eq '${value}'`)();
  //       console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
  //       const normalizeItem2 = async (item: any) => ({
  //         Log: item?.Log || '', // Replace with appropriate mappings
  //         CurrentUser: item?.Folderdetail?.CurrentUser || '',
  //         Remark: item?.Remark || '',
  //         LogHistory: item?.LogHistory || '',
  //         // ProcessName:  item?.Folderdetail?.Processname,
  //         RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //         FileUID: {
  //           FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
  //           SiteName: item?.FolderMeta?.SiteName || '',
  //           DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
  //           // FileName: item?.FolderMeta?.FolderName || '',
  //           FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
  //           RequestNo: item?.Folderdetail?.RequestNo || '',
  //           Status: item?.Folderdetail?.Status || '',
  //           FolderPath: item?.Folderdetail?.FolderPath || '',
  //           // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
  //           // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //           Created: item?.Created || '',
  //           ApproveAction: item?.ApproveAction || '',
  //           Processname: item?.Folderdetail?.Processname
  //         },
  //         MasterApproval: {
  //           ApprovalType: item?.ApprovalType || '',
  //           Level: item?.Level || '',
  //           DocumentLibraryName: item?.DocumentLibraryName || ''
  //         }
  //       });
  //       // const normalizeItem3 = Item2.map(normalizeItem2);
  //       const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
  //       console.log("normalizeItem3", normalizeItem3);
  //       console.log(approvalData, "approvalData 666");
  //       const CombinedItems = [...approvalData, ...normalizeItem3];
  //       console.log(CombinedItems, "CombinedItems")
  //       setMylistdata(CombinedItems);
  //       // setMylistdata(updatedItems);


  //       return arr = CombinedItems
  //     }
  //   } catch (error) {
  //     console.error("Error fetching list items:", error);
  //   }
  // };
  // Add this function near your other data fetching functions
  // const getProjectWorkflowApprovals = async (status: string, actingfor?: any) => {
  //   try {
  //     const currentUserEmail = actingfor || currentUserEmailRef.current;

  //     // Fetch from ProjectApprovals list (adjust the list name and fields as needed)
  //     const items = await sp.web.lists.getByTitle('ProjectApprovals').items.select(
  //       "*",
  //       "ID",
  //       "Status",
  //       "RequestedBy/Title",
  //       "RequestedBy/EMail",
  //       "Created",
  //     )
  //       .expand("RequestedBy")
  //       .filter(`Status eq '${status}'`).orderBy("Created", false).getAll();

  //     console.log(items, "ProjectApprovals List");

  //     // Format the data to match your table structure
  //     const formattedItems = items.map((item: any) => ({
  //       Id: item.ID,
  //       RequestID: `PROJ-${item.ID}`,
  //       ProcessName: "Project Workflow",
  //       Status: item.Status,
  //       Requester: {
  //         Title: item.RequestedBy?.Title || "",
  //         EMail: item.RequestedBy?.EMail || ""
  //       },
  //       Author: {
  //         Title: item.RequestedBy?.Title || "", 
  //       },
  //       Created: item.Created,
  //       // Add any other project-specific fields you need
  //     }));

  //     return formattedItems;
  //   } catch (error) {
  //     console.error("Error fetching Project Workflow items:", error);
  //     return [];
  //   }
  // };


  const getProjectWorkflowApprovals = async (status: string, actingfor?: any) => {
    try {
      const currentUserEmail = actingfor || currentUserEmailRef.current;

      // Get current user ID for filtering
      const currentUser = await sp.web.currentUser();
      const currentUserId = currentUser.Id;

      // Fetch main approval items with the same fields and expands as your example
      const items = await sp.web.lists.getByTitle("ProjectApprovals").items
        .select(
          "*",
          "DeliverablesDetailsId/Deliverables",
          "DeliverablesDetailsId/DocNumber",
          "DeliverablesDetailsId/Organization",
          "DeliverablesDetailsId/Area",
          "DeliverablesDetailsId/ID",
          "ProjectCreationListID/ProjectName",
          "ProjectCreationListID/ID",
          "AssignedTo/ID",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "Author/Title",
          "Author/EMail"
        )
        .expand("DeliverablesDetailsId", "ProjectCreationListID", "AssignedTo", "Author")
        .filter(`AssignedTo/ID eq ${currentUserId} and Status eq '${status}' and ApproverRole ne 'Vendor'`)
        .orderBy("Created", false)();

      console.log(items, "ProjectApprovals List with current user filter");

      // Transform items with additional data from ProjectCreationList (similar to your example)
      const approvals = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            let creationItem = null;

            // Fetch additional project creation data if ProjectCreationListID exists
            if (item.ProjectCreationListID?.ID) {
              creationItem = await sp.web.lists.getByTitle("ProjectCreationList").items.getById(item.ProjectCreationListID.ID)
                .select(
                  "*",
                  "PreparedBy/Title",
                  "ProjectType/ProjectType",
                  "ProjectType/Id",
                  "ClientName"
                )
                .expand("ProjectType", "PreparedBy")();
            }

            // Format the data to match your table structure
            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ApprovalTitle: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || ""
              },
              Author: {
                Title: item.Author?.Title || "",
              },
              Created: item.Created,
              InitiatedBy: creationItem?.AuthorId,

              // Project-specific fields from your example
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: creationItem?.ProjectType?.ProjectType || "",
              ClientName: creationItem?.ClientName || "",
              PreparedBy: creationItem?.PreparedBy?.Title || "",
              Deliverable: item.DeliverablesDetailsId?.Deliverables || "",
              Area: item.DeliverablesDetailsId?.Area || "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: item.DeliverablesDetailsId?.Organization || "",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created ? new Date(item.Created).toLocaleDateString('en-GB') : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",

              // For redirection or additional actions
              RedirectionLink: item.RedirectionLink || "", // Add if you have this field

              // Additional fields that might be useful for filtering
              SNo: index + 1
            };
          } catch (error) {
            console.error(`Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`, error);

            // Return fallback data if creationItem fetch fails
            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ApprovalTitle: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || ""
              },
              Author: {
                Title: item.Author?.Title || "",
              },
              Created: item.Created,
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: "",
              ClientName: "",
              PreparedBy: "",
              Deliverable: item.DeliverablesDetailsId?.Deliverables || "",
              Area: item.DeliverablesDetailsId?.Area || "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: item.DeliverablesDetailsId?.Organization || "",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created ? new Date(item.Created).toLocaleDateString('en-GB') : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",
              RedirectionLink: item.RedirectionLink || "",
              SNo: index + 1
            };
          }
        })
      );

      console.log("Transformed Project Approvals:", approvals);
      return approvals;

    } catch (error) {
      console.error("Error fetching Project Workflow items:", error);
      return [];
    }
  };

  // Add these functions
  const addNewProjectApprovalRow = (item?: any) => {
    const newRow: ApprovalHierarchyItem = {
      level: `Level ${projectHierarchy.length + 1}`,
      approverRole: item?.ApproverRole || '',
      approver: item?.AssignedTo?.Title || '',
      approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
      approvalCriteria: item?.ApprovalCriteria || 'Anyone',
      assignedTo: item?.AssignedTo ? [item.AssignedTo] : []
    };
    setProjectHierarchy(prev => [...prev, newRow]);
  };

  const deleteProjectApprovalRow = (index: number) => {
    setProjectHierarchy(prev => {
      const updatedHierarchy = prev.filter((_, i) => i !== index);
      // Renumber the levels sequentially
      return updatedHierarchy.map((row, i) => ({
        ...row,
        level: `Level ${i + 1}`
      }));
    });
  };

  const updateProjectApprovalRow = (index: number, field: keyof ApprovalHierarchyItem, value: string) => {
    setProjectHierarchy(prev => prev.map((row, i) =>
      i === index ? { ...row, [field]: value } : row
    ));
  };

  const handleProjectApproverChange = (index: number, selectedOptions: any) => {
    setProjectHierarchy(prev => prev.map((row, i) => {
      if (i === index) {
        const selectedUsers = selectedOptions || [];
        const approverNames = selectedUsers.map((user: any) => user.label).join(', ');
        const assignedToArray = selectedUsers.map((user: any) => ({
          ID: parseInt(user.value),
          Title: user.label,
          EMail: user.email
        }));

        return {
          ...row,
          approver: approverNames,
          assignedTo: assignedToArray
        };
      }
      return row;
    }));
  };

  const getApprovalmasterTasklist = async (value: any, actingfor?: any) => {
    // alert(`Status value is ${value} is acting for ${actingfor} in DMS`)

    try {
      // Retrieve current user email
      const currentUserEmail = currentUserEmailRef.current;

      // Fetch the ARGDelegateList items where the current user is in the ActingFor column
      const today = new Date().toISOString(); // Get today's date in YYYY-MM-DD format
      // console.log("today", today);

      // const delegateListItems = await sp.web.lists.getByTitle('ARGDelegateList').items.select(
      //   "DelegateName/EMail",
      //   "ActingFor/EMail",
      //   "ActingFor/Title",
      //   "StartDate",
      //   "EndDate",
      //   "Status"
      // )
      // .expand("DelegateName", "ActingFor")
      // .filter(`ActingFor/EMail eq '${currentUserEmail}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)();

      // console.log("delegateListItems", delegateListItems);

      // const additionalFilters = delegateListItems.map((item:any) => `CurrentUser eq '${item.DelegateName.EMail}'`).join(' or ');
      // const combinedFilters = `CurrentUser eq '${currentUserEmail}'${additionalFilters ? ` or (${additionalFilters})` : ''} and FileUID/Status eq '${value}'`;

      // // Fetch items from DMSFileApprovalTaskList based on the combined filters
      // const items2 = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
      //   "Log", "CurrentUser", "Remark", "LogHistory", "FileUID/FileUID",
      //   "FileUID/SiteName", "FileUID/DocumentLibraryName", "FileUID/FileName",
      //   "FileUID/RequestNo", "FileUID/Processname", "FileUID/Status",
      //   "FileUID/FolderPath", "FileUID/RequestedBy", "FileUID/Created",
      //   "FileUID/ApproveAction", "MasterApproval/ApprovalType", "MasterApproval/Level",
      //   "MasterApproval/DocumentLibraryName"
      // )
      // .expand("FileUID", "MasterApproval")
      // .filter(combinedFilters)
      // .orderBy("Created", false)
      // .getAll();

      // console.log(items2, "DMSFileApprovalTaskList");



      let arr = [];
      let approvalData: any[] = [];
      if (!actingfor) {
        const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
          "Log", "CurrentUser", "Remark"
          , "LogHistory"
          , "FileUID/FileUID"
          , "FileUID/SiteName"
          , "FileUID/DocumentLibraryName"
          , "FileUID/FileName"
          , "FileUID/RequestNo"
          , "FileUID/Processname"
          //  ,"FileUID/FilePreviewUrl" 
          , "FileUID/Status"
          , "FileUID/FolderPath"
          , "FileUID/RequestedBy"
          , "FileUID/Created"
          , "FileUID/ApproveAction"
          , "MasterApproval/ApprovalType"
          , "MasterApproval/Level"
          , "MasterApproval/DocumentLibraryName"

        )
          .expand("FileUID", "MasterApproval")
          .filter(`CurrentUser eq '${currentUserEmailRef.current}' and FileUID/Status eq '${value}'`).orderBy("Created", false).getAll();
        console.log(items, "DMSFileApprovalTaskList");
        items.map((item) => {
          if (item.CurrentUser !== currentUserEmailRef.current) {
            arr.push(item)
            // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
          }

        });
        const updatedItems = await Promise.all(items.map(async (item) => {
          const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
          return { ...item, RequestedByTitle: requestedbyuserTitle };
        }));
        approvalData = updatedItems
        setMylistdata(updatedItems);
      }
      if (actingfor !== "" && actingfor !== undefined) {
        const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
          "Log", "CurrentUser", "Remark"
          , "LogHistory"
          , "FileUID/FileUID"
          , "FileUID/SiteName"
          , "FileUID/DocumentLibraryName"
          , "FileUID/FileName"
          , "FileUID/RequestNo"
          , "FileUID/Processname"
          //  ,"FileUID/FilePreviewUrl" 
          , "FileUID/Status"
          , "FileUID/FolderPath"
          , "FileUID/RequestedBy"
          , "FileUID/Created"
          , "FileUID/ApproveAction"
          , "MasterApproval/ApprovalType"
          , "MasterApproval/Level"
          , "MasterApproval/DocumentLibraryName"

        )
          .expand("FileUID", "MasterApproval")
          .filter(`CurrentUser eq '${actingfor}' and FileUID/Status eq '${value}'`).orderBy("Created", false).getAll();
        console.log(items, "DMSFileApprovalTaskList");
        items.map((item) => {
          if (item.CurrentUser !== currentUserEmailRef.current) {
            arr.push(item)
            // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
          }

        });
        const updatedItems = await Promise.all(items.map(async (item) => {
          const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
          return { ...item, RequestedByTitle: requestedbyuserTitle };
        }));
        approvalData = updatedItems;
        setMylistdata(updatedItems);
      }

      // const updatedItems2 = await Promise.all(items2.map(async (item) => {
      //   const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
      //   return { ...item, RequestedByTitle: requestedbyuserTitle };
      // }));
      if (!actingfor) {
        const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
          "*",
          "Folderdetail"
          , "Folderdetail/SiteTitle"
          , "Folderdetail/DocumentLibraryName"
          , "Folderdetail/CurrentUser"
          , "Folderdetail/FolderPath"
          , "Folderdetail/FolderName"
          , "Folderdetail/ParentFolderId"
          , "Folderdetail/Department"
          , "Folderdetail/Devision"
          , "Folderdetail/RequestNo"
          , "FolderMeta"
          , "FolderMeta/SiteName"
          , "FolderMeta/DocumentLibraryName"
          , "FolderMeta/ColumnName",
          "Folderdetail/Processname",
          "Folderdetail/Status",
          "Approver"
        ).expand("Folderdetail", "FolderMeta")
          .filter(`Approver eq '${currentUserEmailRef.current}' and Folderdetail/Status eq '${value}'`)();
        console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
        const normalizeItem2 = async (item: any) => ({
          Log: item?.Log || '', // Replace with appropriate mappings
          CurrentUser: item?.Folderdetail?.CurrentUser || '',
          Remark: item?.Remark || '',
          LogHistory: item?.LogHistory || '',
          // ProcessName:  item?.Folderdetail?.Processname,
          RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
          FileUID: {
            FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
            SiteName: item?.FolderMeta?.SiteName || '',
            DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
            // FileName: item?.FolderMeta?.FolderName || '',
            FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
            RequestNo: item?.Folderdetail?.RequestNo || '',
            Status: item?.Folderdetail?.Status || '',
            FolderPath: item?.Folderdetail?.FolderPath || '',
            // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
            // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
            Created: item?.Created || '',
            ApproveAction: item?.ApproveAction || '',
            Processname: item?.Folderdetail?.Processname
          },
          MasterApproval: {
            ApprovalType: item?.ApprovalType || '',
            Level: item?.Level || '',
            DocumentLibraryName: item?.DocumentLibraryName || ''
          }
        });
        // const normalizeItem3 = Item2.map(normalizeItem2);
        const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
        console.log("normalizeItem3", normalizeItem3);
        console.log(approvalData, "approvalData 666");
        const CombinedItems = [...approvalData, ...normalizeItem3];
        console.log(CombinedItems, "CombinedItems")
        setMylistdata(CombinedItems);
        // setMylistdata(updatedItems);


        // return arr = CombinedItems
      }
      if (actingfor !== "" && actingfor !== undefined) {
        const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
          "*",
          "Folderdetail"
          , "Folderdetail/SiteTitle"
          , "Folderdetail/DocumentLibraryName"
          , "Folderdetail/CurrentUser"
          , "Folderdetail/FolderPath"
          , "Folderdetail/FolderName"
          , "Folderdetail/ParentFolderId"
          , "Folderdetail/Department"
          , "Folderdetail/Devision"
          , "Folderdetail/RequestNo"
          , "FolderMeta"
          , "FolderMeta/SiteName"
          , "FolderMeta/DocumentLibraryName"
          , "FolderMeta/ColumnName",
          "Folderdetail/Processname",
          "Folderdetail/Status",
          "Approver"
        ).expand("Folderdetail", "FolderMeta")
          .filter(`Approver eq '${actingfor}' and Folderdetail/Status eq '${value}'`)();
        console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
        const normalizeItem2 = async (item: any) => ({
          Log: item?.Log || '', // Replace with appropriate mappings
          CurrentUser: item?.Folderdetail?.CurrentUser || '',
          Remark: item?.Remark || '',
          LogHistory: item?.LogHistory || '',
          // ProcessName:  item?.Folderdetail?.Processname,
          RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
          FileUID: {
            FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
            SiteName: item?.FolderMeta?.SiteName || '',
            DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
            // FileName: item?.FolderMeta?.FolderName || '',
            FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
            RequestNo: item?.Folderdetail?.RequestNo || '',
            Status: item?.Folderdetail?.Status || '',
            FolderPath: item?.Folderdetail?.FolderPath || '',
            // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
            // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
            Created: item?.Created || '',
            ApproveAction: item?.ApproveAction || '',
            Processname: item?.Folderdetail?.Processname
          },
          MasterApproval: {
            ApprovalType: item?.ApprovalType || '',
            Level: item?.Level || '',
            DocumentLibraryName: item?.DocumentLibraryName || ''
          }
        });
        // const normalizeItem3 = Item2.map(normalizeItem2);
        const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
        console.log("normalizeItem3", normalizeItem3);
        console.log(approvalData, "approvalData 666");
        const CombinedItems = [...approvalData, ...normalizeItem3];
        console.log(CombinedItems, "CombinedItems")
        setMylistdata(CombinedItems);
        // setMylistdata(updatedItems);


        // return arr = CombinedItems
      }
    } catch (error) {
      console.error("Error fetching list items:", error);
    }
  };
  console.log(Mylistdata, "Mylistdata");
  const currentUserEmailRef = useRef("");
  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getApprovalmasterTasklist('Pending', '');
    myActingfordata()
  };
  const fetchUsers = async () => {
    try {
      const siteUsers = await sp.web.siteUsers();
      const userOptions: UserOption[] = siteUsers
        .filter((user: any) => user.Email && user.Title) // Filter out users without email or title
        .map((user: any) => ({
          value: user.Id.toString(),
          label: user.Title,
          email: user.Email
        }));
      setUsers(userOptions);
    } catch (error) {
      console.error("Error fetching users:", error);
      // Fallback to some default users if fetch fails
      setUsers([
        { value: '1', label: 'Current User', email: props.context.pageContext.user.email },
        { value: '2', label: 'Admin User', email: 'admin@contoso.com' }
      ]);
    }
  };
  React.useEffect(() => {
    getCurrrentuser();
    fetchUsers();
  }, []);





  const truncateText = (text: string, maxLength?: any) => {
    if (text) {
      return text.length > maxLength
        ? text.substring(0, maxLength) + "..."
        : text;
    }
  };

  const getTaskItemsbyID = async (e: any, itemid: any, ProcessName: string) => {
    // currentItemID = itemid
    currentItemID = itemid;
    setActiveComponent("Approval Action");
    setFolderActionOrFileAction(ProcessName);
    console.log("itemid", itemid);
    // const items = await sp.web.lists
    //   .getByTitle("DMSFileApprovalTaskList")
    //   .items.select("CurrentUser", "FileUID/FileUID", "Log")
    //   .expand("FileUID")
    //   .filter(`FileUID/RequestNo eq '${itemid}'`)();
    // console.log(items, "items");
  };
  const getTaskItemsbyID2 = async (e: any, itemid: any) => {
    // alert("Folder")
    // currentItemID = itemid
    currentItemID = itemid
    setActiveComponent('DMS Folder Approval')
    console.log("itemid", itemid)
    // const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser" , "FileUID/FileUID" , "Log").expand("FileUID").filter(`FileUID/RequestNo eq '${itemid}'`)();
    //    console.log(items , "items")
  }
  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };

  const [sortConfig, setSortConfig] = React.useState({
    key: "",

    direction: "ascending",
  });

  const [formData, setFormData] = React.useState({
    Remark: "",
  });
  const handleCancel = () => {
    window.location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;
  };
  //#region OnchangeData

  const onChange = (name: string, value: string) => {
    debugger;

    setFormData((prevData) => ({
      ...prevData,

      [name]: value,
    }));
  };

  //#endregion

  // const [formData, setFormData] = React.useState({

  //   topic: "",

  //   category: "",

  //   entity: "",

  //   Type: "",

  //   GroupType: "",

  //   description: "",

  //   overview: "",

  //   FeaturedAnnouncement: false,

  // });

  const [approveData, setApproveData] = useState([]);

  const [isActivedata, setisActivedata] = useState(false);

  const [DiscussionData, setDiscussion] = useState([]);

  const [CategoryData, setCategoryData] = React.useState([]);
  const [Statusvalue, SetStatusvalue] = useState("Pending");
  const [showModal, setShowModal] = React.useState(false);
  const [StatusChange, setStatusChange] = React.useState(false);
  const [showDocTable, setShowDocTable] = React.useState(false);

  const [showImgModal, setShowImgTable] = React.useState(false);

  const [showBannerModal, setShowBannerTable] = React.useState(false);

  const [currentUser, setCurrentUser] = React.useState(null);

  const [editForm, setEditForm] = React.useState(false);

  const [richTextValues, setRichTextValues] = React.useState<{
    [key: string]: string;
  }>({});

  const [documentControllerId, setDocumentControllerId] = React.useState<number | null>(null);
  const [dccId, setDccId] = React.useState<number | null>(null);


  //const [activeTab, setActiveTab] = useState("home1");
  const [activeTab, setActiveTab] = useState("Automation");
  const handleTabClick = async (tab: React.SetStateAction<string>) => {

    setActiveTab(tab);
    debugger
    setMyApprovalsData([]);
    setLoading(true);
    // console.log(
    //   "tab",
    //   tab,
    //   myApprovalsDataAutomation,
    //   myApprovalsDataAll,
    //   myApprovalsData
    // );

    // if (tab == "Intranet") {
    //   setMyApprovalsData(myApprovalsDataAll);
    // } else if (tab == "DMS") {
    //   setMyApprovalsData(Mylistdata);
    // } else if (tab == "Automation") {
    //   //ApiCall("Pending");
    //   setMyApprovalsData(myApprovalsDataAutomation);
    //   //setMyApprovalsDataAutomation(myApprovalsDataAutomation);
    // }
    let MyApprovaldata: any = [];
    let Automationdata: any = [];
    // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
    let MyDMSAPPROVALDATA: any = [];
    let ProjectWorkflowData: any = []; // NEW
    if (actingforuseremail === undefined || actingforuseremail === null || actingforuseremail === "") {
      MyApprovaldata = await getMyApproval(sp, Statusvalue);
      Automationdata = await getApprovalListsData(sp, Statusvalue);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      MyDMSAPPROVALDATA = await getApprovalmasterTasklist(Statusvalue)
      ProjectWorkflowData = await getProjectWorkflowApprovals(Statusvalue); // NEW
    } else {
      MyApprovaldata = await getMyApproval(sp, Statusvalue, actingforuseremail);
      Automationdata = await getApprovalListsData(sp, Statusvalue, actingforuseremail);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      MyDMSAPPROVALDATA = await getApprovalmasterTasklist(Statusvalue, actingforuseremail)
      ProjectWorkflowData = await getProjectWorkflowApprovals(Statusvalue, actingforuseremail); // NEW
    }
    console.log("MyDMSAPPROVALDATA", MyDMSAPPROVALDATA)
    setMyApprovalsDataAll(MyApprovaldata);
    setMyApprovalsDataAutomation(Automationdata);
    setProjectWorkflowdata(ProjectWorkflowData); // NEW
    if (tab == "Intranet") {
      setMyApprovalsData(MyApprovaldata);
      if (MyApprovaldata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }

    } else if (tab == "DMS") {
      // alert(value)
      // setMyApprovalsData(MyDMSAPPROVALDATA);
      setMyApprovalsData(Mylistdata);
      if (Mylistdata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
    } else if (tab == "ProjectWorkflow") { // NEW SECTION
      setMyApprovalsData(ProjectWorkflowData);
      if (ProjectWorkflowData.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
    } else if (tab == "Automation") {
      setMyApprovalsData(Automationdata.sort((a: any, b: any) => b.Created - a.Created));
      if (Automationdata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
      console.log("Automationdata", Automationdata);
    }

  };

  const getProjectConfiguration = async (): Promise<{ documentControllerId: number | null; dccId: number | null }> => {
    try {
      const items = await sp.web.lists
        .getByTitle("ProjectConfiguration")
        .items.select(
          "*",
          "DocumentController/ID",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DCC/ID",
          "DCC/Title",
          "DCC/EMail"
        )
        .expand("DocumentController", "DCC")
        .orderBy("Created", false)();

      if (items.length > 0) {
        const documentControllerId = items[0]?.DocumentControllerId || null;
        const dccId = items[0]?.DCCId || null;

        console.log("Document Controller ID:", documentControllerId);
        console.log("DCC ID:", dccId);

        return { documentControllerId, dccId };
      }
      return { documentControllerId: null, dccId: null };
    } catch (error) {
      console.error("Error fetching Project Configuration:", error);
      return { documentControllerId: null, dccId: null };
    }
  };
  React.useEffect(() => {
    const initializeData = async () => {
      const { documentControllerId, dccId } = await getProjectConfiguration();
      setDocumentControllerId(documentControllerId);
      setDccId(dccId);
    };

    initializeData();
  }, []);


  // Add this function near your other action handlers
  const handleProjectWorkflowAction = async (e: any, item: any, mode: string) => {
    e.preventDefault();

    // For both view and approval modes, just show the form details
    await handleProjectViewClick(item);
  };

  // Handle project view click - just sets the data and shows form
  const handleProjectViewClick = async (task: any) => {
    setSelectedProjectTask(task);
    setShowProjectForm(true);
    setProjectNeedsFurtherApproval(task.DoYouNeedApproval || "Select");
    setProjectRemarks(task.Remarks || "");

    // You can keep document fetching if needed for display
    if (task.DeliverableId) {
      const docInfo = await fetchDocumentForDeliverable(task.DeliverableId);
      setProjectDocumentInfo(docInfo);
      // You might want to set document info state if needed for display
    }
    // Initialize projectHierarchy if needed
    if (task.ProjectId && task.DeliverableId) {
      await getProjectApprovalHierarchy(task.ProjectId, task.DeliverableId, task.DocType);
    } else {
      addNewProjectApprovalRow();
    }
  };

  const getProjectApprovalHierarchy = async (projectId: number, deliverableId: number, documentType: string) => {
    try {
      const items = await sp.web.lists.getByTitle("ApprovalHierarchy").items
        .select("*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail")
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(`ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`)
        .orderBy("SerialNumber", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map((item: any, index: number) => {
          // AssignedTo will be an array (for multi-select people fields)
          const assignedArray = Array.isArray(item.AssignedTo) ? item.AssignedTo : [];

          return {
            id: item.Id,
            level: `Level ${index + 1}`,
            approverRole: item.ApproverRole || '',
            approver: assignedArray.map((a: any) => a.Title).join(", "),
            approvers: assignedArray.map((a: any) => a.Title),
            approvalCriteria: item.ApprovalCriteria || 'Anyone',
            serialNumber: item.SerialNumber,
            assignedTo: assignedArray.map((a: any) => ({
              ID: a.ID,
              Title: a.Title,
              EMail: a.EMail
            }))
          };
        });

        setProjectHierarchy(hierarchyItems);
      } else {
        await getProjectWorkflowConfiguration(documentType);
      }
    } catch (error) {
      console.error("Error fetching approval projectHierarchy:", error);
      await getProjectWorkflowConfiguration(documentType);
    }
  };

  // Get Project Workflow Configuration
  const getProjectWorkflowConfiguration = async (docType: string) => {
    try {
      const items = await sp.web.lists.getByTitle("ProjectWorflowConfiguration").items
        .select("*,DocumentType/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail")
        .expand("DocumentType,AssignedTo")
        .filter(`DocumentType/DocumentType eq '${docType}'`)
        .orderBy("ID", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map((item: any, index: number) => ({
          level: `Level ${index + 1}`,
          approverRole: item.Role || '',
          approver: item.AssignedTo?.Title || '',
          approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
          approvalCriteria: item.ApprovalCriteria || 'Anyone',
          assignedTo: item.AssignedTo ? [item.AssignedTo] : []
        }));
        setProjectHierarchy(hierarchyItems);
      } else {
        // Add one default row if no configuration found
        addNewProjectApprovalRow();
      }
    } catch (error) {
      console.error("Error fetching workflow configuration:", error);
      addNewProjectApprovalRow();
    }
  };

  // Update the handleProjectBackClick to reset all states
  const handleProjectBackClick = () => {
    setSelectedProjectTask(null);
    setShowProjectForm(false);
    setProjectHierarchy([]);
    setProjectNeedsFurtherApproval("Select");
    setProjectRemarks("");
    setProjectDocumentInfo(null);
  };

  // Simple project document open function (if needed)
  // Handle project document open with proper URL construction
  const handleProjectOpenDocument = () => {
    if (!projectDocumentInfo) {
      console.error('No document available');
      alert('No document available to open.');
      return;
    }

    try {
      let documentUrl = '';

      // Priority 1: Use SharedLink if available
      if (projectDocumentInfo.sharedLink) {
        documentUrl = projectDocumentInfo.sharedLink;
        console.log('Opening document using SharedLink:', documentUrl);
      }
      // Priority 2: Use the document URL to construct the full URL
      else if (projectDocumentInfo.documentUrl) {
        const siteUrl = props.siteUrl; // Use your siteUrl from props
        documentUrl = `${siteUrl}${projectDocumentInfo.documentUrl}`;
        console.log('Opening document using ServerRelativeUrl:', documentUrl);
      }
      // Priority 3: Use fileRef if available
      else if (projectDocumentInfo.fileRef) {
        const siteUrl = props.siteUrl;
        documentUrl = `${siteUrl}${projectDocumentInfo.fileRef}`;
        console.log('Opening document using FileRef:', documentUrl);
      }
      else {
        throw new Error('No valid document URL found');
      }

      // Open the document in a new tab
      window.open(documentUrl, '_blank', 'noopener,noreferrer');

      console.log(`Opened document: ${projectDocumentInfo.fileName || projectDocumentInfo.fileLeafRef}`);

    } catch (error) {
      console.error('Error opening document:', error);
      alert('Error opening document. Please try again or contact administrator.');
    }
  };

  // Function to fetch document for a specific deliverable
  const fetchDocumentForDeliverable = async (deliverableId: number) => {
    try {
      const documents = await sp.web.lists.getByTitle("DeliverablesDocument").items
        .select("*,File/ServerRelativeUrl,FileLeafRef,FileRef,Author/ID,Author/Title,SharedLink")
        .expand("File", "Author")
        .filter(`DeliverablesDetailsId eq ${deliverableId}`)
        .orderBy("ID", false)
        .top(1)(); // Get the latest document

      if (documents.length > 0) {
        const latestDoc = documents[0];
        return {
          documentUrl: latestDoc.File.ServerRelativeUrl,
          fileName: latestDoc.FileLeafRef,
          fileLeafRef: latestDoc.FileLeafRef,
          fileRef: latestDoc.FileRef,
          sharedLink: latestDoc.SharedLink // Get the SharedLink column value
        };
      } else {
        console.warn(`No documents found for deliverable ID: ${deliverableId}`);
        return null;
      }
    } catch (error) {
      console.error(`Error fetching documents for deliverable ${deliverableId}:`, error);
      return null;
    }
  };

  // Project Approval Code starts here ----------------
  const handleProjectApprovalAction = async (status: "Approved" | "Rejected" | "Rework") => {
    if (!selectedProjectTask) return;

    try {
      // Update the current approval status
      await updateCurrentApprovalStatus(status);
      console.log("Status updated successfully, now checking status type:", status);

      let actionExecuted = false; // ✅ Flag to control alert

      if (status === "Approved") {
        await handleApprovedAction();
        actionExecuted = true;
      } else if (status === "Rework") {
        await handleReworkAction();
        actionExecuted = true;
      } else if (status === "Rejected") {
        await handleRejectedAction();
        actionExecuted = true;
      }

      if (actionExecuted) { // ✅ Only show alert if something executed
        alert(`${status} successfully!`);
        handleProjectBackClick();
        window.location.reload(); // optional refresh
      } else {
        console.warn("⚠️ No matching action found for status:", status);
      }
    } catch (error) {
      console.error("Error in approval action:", error);
      alert("Something went wrong while processing approval.");
    }
  };
  // Update current approval status
  const updateCurrentApprovalStatus = async (status: string) => {
    if (!selectedProjectTask) return;

    try {
      const data: any = {
        Status: status,
        Remarks: projectRemarks,
        ApprovalDate: new Date(),
      };

      if (projectNeedsFurtherApproval !== "Select") {
        data.Doyouneedapproval = projectNeedsFurtherApproval;
      }

      if (status === "Rework" && selectedProjectTask.ApprovalRole === "Document Controller") {
        data.OutgoingDate = new Date();
      }

      await sp.web.lists.getByTitle("ProjectApprovals").items.getById(selectedProjectTask.Id).update(data);
      console.log("Status Updated");
    } catch (error) {
      console.error("Error in updateCurrentApprovalStatus:", error);
      throw error; // Re-throw to be caught in main handler
    }
  };

  // Handle Approved action
  const handleApprovedAction = async () => {
    if (!selectedProjectTask) return;

    const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);
    console.log('Array?', Array.isArray(approvals));
    await addMultipleApproval("Approved", approvals);
  };

  // Get all project approvals
  const getAllProjectApproval = async (projectId: number, deliverableId: number) => {
    try {
      const items = await sp.web.lists.getByTitle("ProjectApprovals").items
        .select("*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title")
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(`ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`)
        .orderBy("Created", false)();

      // setProjectApprovalArr(items);
      return items;
    } catch (error) {
      console.error("Error fetching project approvals:", error);
      return [];
    }
  };

  // Add Multiple Approval logic
  const addMultipleApproval = async (statusUpdate: string, allApprovals: any[]) => {
    if (!selectedProjectTask) return;
    console.log('Array?', Array.isArray(allApprovals), allApprovals.length);
    try {
      if (statusUpdate === "Approved") {
        const currentRole = selectedProjectTask.ApprovalRole;
        const currentCriteria = selectedProjectTask.ApprovalCriteria;

        if (currentRole === "Document Controller" && projectNeedsFurtherApproval === "Yes") {
          // Create projectHierarchy and approvals for multiple approvers
          for (let index = 0; index < projectHierarchy.length; index++) {
            const row = projectHierarchy[index];
            const approverAssignOwner = row.assignedTo?.map(user => user.ID) || [];

            // Update or create projectHierarchy
            const rowData = {
              ProjectCreationListIDId: selectedProjectTask.ProjectId,
              DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
              // AssignedToId: { 'results': approverAssignOwner },
              AssignedToId: approverAssignOwner,
              RequestedById: props.context.pageContext.legacyPageContext.userId,
              RequestedDate: new Date(),
              DocumentType: selectedProjectTask.DocType,
              ApproverRole: row.approverRole,
              ApprovalCriteria: row.approvalCriteria,
              Status: 'Pending',
              Doyouneedapproval: projectNeedsFurtherApproval,
              Level: row.level,
              SerialNumber: row.serialNumber || index + 1,
            };

            if (row.id) {
              await sp.web.lists.getByTitle("ApprovalHierarchy").items.getById(row.id).update(rowData);
            } else {
              await sp.web.lists.getByTitle("ApprovalHierarchy").items.add(rowData);
            }

            // Create approvals for Level 1
            if (row.level === "Level 1") {
              for (const userId of approverAssignOwner) {
                const approvalData = {
                  ProjectCreationListIDId: selectedProjectTask.ProjectId,
                  DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
                  AssignedToId: userId,
                  RequestedById: props.context.pageContext.legacyPageContext.userId,
                  RequestedDate: new Date(),
                  DocumentType: selectedProjectTask.DocType,
                  ApproverRole: row.approverRole,
                  ApprovalCriteria: row.approvalCriteria,
                  Doyouneedapproval: projectNeedsFurtherApproval,
                  Status: 'Pending',
                  Level: row.level,
                  SerialNumber: row.serialNumber || index + 1,
                  RevisionNumber: selectedProjectTask.RevisionNumber,
                };
                await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
              }
            }
          }
        } else if (currentRole === "Document Controller" && projectNeedsFurtherApproval === "No") {
          // Create DCC approval
          const approvalData = {
            ProjectCreationListIDId: selectedProjectTask.ProjectId,
            DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
            AssignedToId: dccId,
            RequestedById: props.context.pageContext.legacyPageContext.userId,
            RequestedDate: new Date(),
            DocumentType: selectedProjectTask.DocType,
            ApproverRole: "DCC",
            ApprovalCriteria: "Anyone",
            Status: 'Pending',
            Level: `Level ${selectedProjectTask.ApprovalSN + 1}`,
            SerialNumber: selectedProjectTask.ApprovalSN + 1,
            RevisionNumber: selectedProjectTask.RevisionNumber,
          };
          await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
          await updateDeliverableStatus(statusUpdate);
        } else {
          // Regular approval flow
          const currentApprovals = allApprovals.filter(x =>
            x.ApproverRole === currentRole &&
            x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
            x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
          );

          const allApproved = currentApprovals.length > 0 &&
            currentApprovals.every(item => item.Status === "Approved" || item.Status === "Rework");

          // Update current approvals if criteria = "Anyone"
          if (currentCriteria === "Anyone" && allApprovals.length > 0) {
            for (const approvalItem of allApprovals) {
              if (approvalItem.Id && approvalItem.ApproverRole === currentRole) {
                const updateData = {
                  Status: statusUpdate,
                  ApprovalDate: new Date(),
                };
                await sp.web.lists.getByTitle("ProjectApprovals").items.getById(approvalItem.Id).update(updateData);
              }
            }
            await updateApprovalHierarchyStatus(allApprovals);
          }

          const shouldProceed = (currentCriteria === "Everyone" && allApproved) ||
            (currentCriteria === "Anyone");

          if (shouldProceed) {
            await handleNextApprovers();
          }
        }
      }
    } catch (error) {
      console.error("Error in addMultipleApproval:", error);
      throw error;
    }
  };

  const updateDeliverableStatus = async (status: string) => {
    if (!selectedProjectTask) return;

    await sp.web.lists.getByTitle("DeliverablesDetails")
      .items.getById(selectedProjectTask.DeliverableId!)
      .update({ Status: status });
  };

  const updateApprovalHierarchyStatus = async (allApprovals: any[]) => {
    if (!selectedProjectTask) return;

    try {
      const currentRole = selectedProjectTask.ApprovalRole;
      const currentCriteria = selectedProjectTask.ApprovalCriteria;

      const currentApprovals = allApprovals.filter(x =>
        x.ApproverRole === currentRole &&
        x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
        x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
      );

      if (currentApprovals.length === 0) return;

      const allApproved = currentApprovals.every(item => item.Status === "Approved");
      let newStatus = "";

      if (currentCriteria === "Anyone") {
        newStatus = "Approved";
      } else if (currentCriteria === "Everyone") {
        newStatus = allApproved ? "Approved" : "In Progress";
      }

      const currentHierarchy = projectHierarchy.find(x =>
        x.approverRole === currentRole &&
        x.id &&
        x.id === selectedProjectTask!.deliverableId
      );

      if (currentHierarchy && currentHierarchy.id && newStatus !== "") {
        await sp.web.lists.getByTitle("ApprovalHierarchy")
          .items.getById(currentHierarchy.id)
          .update({ Status: newStatus });
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatus:", error);
    }
  };

  // Handle next approvers
  const handleNextApprovers = async () => {
    if (!selectedProjectTask) return;

    const currentHierarchy = projectHierarchy.find(x =>
      x.approverRole === selectedProjectTask.ApprovalRole &&
      x.level === selectedProjectTask.CurrentApprovalLevel
    );

    const nextApprovers = currentHierarchy ?
      projectHierarchy.filter(x => x.serialNumber === (currentHierarchy.serialNumber! + 1)) : [];

    if (!nextApprovers || nextApprovers.length === 0) {
      // Final stage
      if (selectedProjectTask.ApprovalRole !== "DCC") {
        const approvalData = {
          ProjectCreationListIDId: selectedProjectTask.ProjectId,
          DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
          AssignedToId: dccId,
          RequestedById: props.context.pageContext.legacyPageContext.userId,
          RequestedDate: new Date(),
          DocumentType: selectedProjectTask.DocType,
          ApproverRole: "DCC",
          ApprovalCriteria: "Anyone",
          Status: 'Pending',
          Level: `Level ${selectedProjectTask.ApprovalSN + 1}`,
          SerialNumber: selectedProjectTask.ApprovalSN + 1,
          RevisionNumber: selectedProjectTask.RevisionNumber,
        };
        await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
        await updateDeliverableStatus("Approved");
      } else if (selectedProjectTask?.ApprovalRole === "DCC" && projectWantsToPublishInDossier !== "Select") {
        await updateDeliverablePublishStatus(selectedProjectTask.DeliverableId!, projectWantsToPublishInDossier);
      }
    } else {
      // Create approvals for next approvers
      // for (const approvalItem of nextApprovers) {
      //     if (approvalItem.assignedTo && approvalItem.assignedTo.length > 0) {
      //         for (const user of approvalItem.assignedTo) {
      //             const approvalData = {
      //                 ProjectCreationListIDId: selectedProjectTask.ProjectId,
      //                 DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      //                 AssignedToId: user[0].ID,
      //                 RequestedById: context.pageContext.legacyPageContext.userId,
      //                 RequestedDate: new Date(),
      //                 DocumentType: selectedProjectTask.DocType,
      //                 ApproverRole: approvalItem.approverRole,
      //                 ApprovalCriteria: approvalItem.approvalCriteria,
      //                 Status: 'Pending',
      //                 Level: approvalItem.level,
      //                 SerialNumber: approvalItem.serialNumber,
      //             };
      //             await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
      //         }
      //     }
      // }

      for (const approvalItem of nextApprovers) {
        if (approvalItem.assignedTo && approvalItem.assignedTo.length > 0) {

          // Flatten in case it's an array of arrays
          const allUsers = ([] as any[]).concat.apply([], approvalItem.assignedTo);

          for (const user of allUsers) {
            const approvalData = {
              ProjectCreationListIDId: selectedProjectTask.ProjectId,
              DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
              AssignedToId: user.ID,
              RequestedById: props.context.pageContext.legacyPageContext.userId,
              RequestedDate: new Date(),
              DocumentType: selectedProjectTask.DocType,
              ApproverRole: approvalItem.approverRole,
              ApprovalCriteria: approvalItem.approvalCriteria,
              Status: 'Pending',
              Level: approvalItem.level,
              SerialNumber: approvalItem.serialNumber,
              RevisionNumber: selectedProjectTask.RevisionNumber,
            };

            await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
          }
        }
      }

    }
  };


  // Handle Rework action
  const handleReworkAction = async () => {
    if (!selectedProjectTask) return;

    if (selectedProjectTask.ApprovalRole !== "Document Controller") {
      const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);

      // Update other pending approvals in current level to Rework
      const pendingApprovals = approvals.filter(x =>
        x.Level === selectedProjectTask?.CurrentApprovalLevel &&
        x.ApproverRole === selectedProjectTask?.ApprovalRole &&
        x.Id !== selectedProjectTask.Id &&
        x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
        x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
      );

      for (const item of pendingApprovals) {
        const data = {
          Status: "Rework",
          Remarks: "Auto Rework",
          ApprovalDate: new Date(),
        };
        await sp.web.lists.getByTitle("ProjectApprovals").items.getById(item.Id).update(data);
      }

      // Create new approval for Document Controller
      await createDocumentControllerApproval();
      await updateApprovalHierarchyStatusForRework("Rework");

    } else {
      // Document Controller Rework logic
      await updateDeliverableStatus("Pending");
      await createVendorApproval();
      // await sendEmail(); // Uncomment if you have email functionality
    }
  };


  const createDocumentControllerApproval = async () => {
    if (!selectedProjectTask) return;

    //const documentControllerId = 6 // Implement this

    const approvalData = {
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: documentControllerId,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Document Controller",
      ApprovalCriteria: "Everyone",
      Doyouneedapproval: "Yes",
      Status: 'Pending',
      Level: "Level 1",
      SerialNumber: 1,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    };
    await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
  };

  const updateApprovalHierarchyStatusForRework = async (status: string) => {
    if (!selectedProjectTask) return;

    try {
      const filterQuery = `ProjectCreationListIDId eq ${selectedProjectTask.ProjectId} and DeliverablesDetailsIdId eq ${selectedProjectTask.DeliverableId} and Level eq '${selectedProjectTask.CurrentApprovalLevel}'`;

      const items = await sp.web.lists.getByTitle("ApprovalHierarchy")
        .items.filter(filterQuery)();

      if (items && items.length > 0) {
        for (const item of items) {
          await sp.web.lists.getByTitle("ApprovalHierarchy")
            .items.getById(item.Id)
            .update({ Status: status });
        }
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatusForRework:", error);
    }
  };

  const createVendorApproval = async () => {
    if (!selectedProjectTask) return;

    const vendorId = 6 // Implement this

    const approvalData = {
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: selectedProjectTask.InitiatedBy,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Vendor",
      ApprovalCriteria: "Anyone",
      Status: 'pending',
      Level: "Level 0",
      SerialNumber: 0,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    };
    await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
  };


  // Handle Rejected action
  const handleRejectedAction = async () => {
    if (!selectedProjectTask) return;

    const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);

    // Update other pending approvals in current level to Rejected
    const pendingApprovals = approvals.filter(x =>
      x.Level === selectedProjectTask?.CurrentApprovalLevel &&
      x.ApproverRole === selectedProjectTask?.ApprovalRole &&
      x.Id !== selectedProjectTask.Id &&
      x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
      x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
    );

    for (const item of pendingApprovals) {
      const data = {
        Status: "Rejected",
        Remarks: "Auto Reject",
        ApprovalDate: new Date(),
      };
      await sp.web.lists.getByTitle("ProjectApprovals").items.getById(item.Id).update(data);
    }

    await updateDeliverableStatus("Rejected");
    await updateApprovalHierarchyStatusForRework("Rejected");
  };


  const updateDeliverablePublishStatus = async (deliverableId: number, publishStatus: string): Promise<void> => {
    try {
      await sp.web.lists.getByTitle("DeliverablesDetails")
        .items.getById(deliverableId)
        .update({
          IsPublished: publishStatus
        });

      console.log("Deliverable publish status updated successfully");
    } catch (error) {
      console.error("Error updating deliverable publish status:", error);
      throw error;
    }
  };
  // Project Approval Code ends here ----------------

  React.useEffect(() => {
    sessionStorage.removeItem("announcementId");

    ApiCall("Pending");
  }, [useHide]);

  // const ApiCall = async (status: string) => {
  //   // if(activeTab == "Intranet"){
  //   setLoading(true);
  //   let MyApprovaldata = await getMyApproval(sp, status);
  //   let Automationdata1 = await getApprovalListsData(sp, status);
  //   let typedata = await getType(sp);
  //   //setMyApprovalsData(MyApprovaldata);
  //   setMyApprovalsDataAll(MyApprovaldata);
  //   //}
  //   //else if(activeTab == "Automation"){
  //   // let Automationdata = Automationdata1.sort((a, b) => {
  //   //   return a.Created === b.Created ? 0 : a.Created ? -1 : 1;
  //   // });
  //   let Automationdata = Automationdata1.sort((a, b) => {
  //     return b.Created - a.Created;
  //   });
  //   setMyApprovalsData(Automationdata);
  //   setMyApprovalsDataAutomation(Automationdata);
  //   setTimeout(() => {
  //     setLoading(false);;
  //   }, 15000);

  //   console.log("Automationdata", Automationdata);
  //   // }
  // };

  // const FilterDiscussionData = async (optionFilter: string) => {

  //   setAnnouncementData(await getDiscussionFilterAll(sp, optionFilter));

  // };

  const ApiCall = async (status: string) => {
    // if(activeTab == "Intranet"){
    setLoading(true);
    let MyApprovaldata = await getMyApproval(sp, status);
    let Automationdata1 = await getApprovalListsData(sp, status);
    let typedata = await getType(sp);
    let ProjectWorkflowData = await getProjectWorkflowApprovals(status); // NEW
    let Automationdata: any;
    //setMyApprovalsData(MyApprovaldata);
    setMyApprovalsDataAll(MyApprovaldata);
    setProjectWorkflowdata(ProjectWorkflowData); // NEW
    //}
    //else if(activeTab == "Automation"){
    // let Automationdata = Automationdata1.sort((a, b) => {
    //   return a.Created === b.Created ? 0 : a.Created ? -1 : 1;
    // });
    if (Automationdata1.length > 0) {
      // Automationdata = Automationdata1.sort((a, b) => {
      //   return b.Created - a.Created;
      // });
      setMyApprovalsData(Automationdata1.sort((a, b) => b.Created - a.Created));
      setMyApprovalsDataAutomation(Automationdata1.sort((a, b) => b.Created - a.Created));
      //setLoading(false)
    } else {
      setMyApprovalsData([]);
      setMyApprovalsDataAutomation([]);
      setLoading(false)
    }
    setLoading(false)
    // setTimeout(() => {
    //   setLoading(false);;
    // }, 15000);

    console.log("Automationdata", Automationdata);
    // }
  };

  const handleStatusChange = async (name: string, value: string, actingfor: any) => {

    setLoading(true);
    setStatusChange(true);
    actingforuseremail = actingfor
    // alert(`Status value is ${value} is acting for ${actingfor}`)
    if (actingforuseremail === undefined || actingforuseremail === null || actingforuseremail === "") {
      // alert("acting for is undefined")
    }

    if (value === "") {
      // Show all records if no type is selected
      console.log("No status selected");
    } else {
      SetStatusvalue(name);
      // Filter records based on the selected type
      let MyApprovaldata = await getMyApproval(sp, value, actingfor);
      let Automationdata = await getApprovalListsData(sp, value, actingfor);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      let MyDMSAPPROVALDATA: any = await getApprovalmasterTasklist(value, actingfor);
      console.log("MyDMSAPPROVALDATA", MyDMSAPPROVALDATA)

      let ProjectWorkflowData = await getProjectWorkflowApprovals(value, actingfor); // NEW
      console.log("ProjectWorkflowData", ProjectWorkflowData);
      setProjectWorkflowdata(ProjectWorkflowData); // NEW
      setMyApprovalsDataAll(MyApprovaldata);
      setMyApprovalsDataAutomation(Automationdata);
      if (activeTab == "Intranet") {
        setMyApprovalsData(MyApprovaldata);
        if (MyApprovaldata.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 5000);
        } else {
          setLoading(false)
        }
      } else if (activeTab == "DMS") {
        // alert(value)
        setMyApprovalsData(MyDMSAPPROVALDATA);
        if (MyDMSAPPROVALDATA.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
      } else if (activeTab == "Automation") {
        setMyApprovalsData(Automationdata.sort((a, b) => b.Created - a.Created));
        if (Automationdata.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
        console.log("Automationdata", Automationdata);
      } else if (activeTab == "ProjectWorkflow") { // NEW
        setMyApprovalsData(ProjectWorkflowData);
        if (ProjectWorkflowData.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
      }
      setStatusChange(false);
      // else if (activeTab == "Automation") {
      //   setMyApprovalsData(null);
      //   setMyApprovalsData(MyDMSAPPROVALDATA);
      //   setMyApprovalsData(Mylistdata);
      //   console.log("Automationdata", Automationdata);
      // }
    }

    setLoading(false);
  };
  const handleFilterChange = (
    e: React.ChangeEvent<HTMLInputElement>,

    field: string
  ) => {
    console.log("eee", e);

    setFilters({
      ...filters,

      [field]: e.target.value,
    });
  };

  const applyFiltersAndSorting = (data: any[]) => {
    // Filter data

    console.log(
      "filters",

      data,

      filters,

      filters.ProcessName,

      filters.RequestID
    );

    const filteredData = data?.filter((item, index) => {
      return (
        (filters.SNo === "" || String(index + 1).includes(filters.SNo)) &&
        // (filters.Title === "" ||
        //   (activeTab == "Intranet" ? item.Title != undefined : item.ApprovalTitle != undefined) &&
        // (activeTab == "Intranet" ? item.Title : item.ApprovalTitle).toLowerCase().includes(filters.Title.toLowerCase()))
        (filters.ProcessName === "" ||
          (item.ProcessName != undefined &&
            item.ProcessName.toLowerCase().includes(
              filters.ProcessName.toLowerCase()
            ))) &&
        (filters.RequestID === "" ||
          (item.RequestID != undefined &&
            item.RequestID.toLowerCase().includes(
              filters.RequestID.toLowerCase()
            ))) &&
        // (filters.ProcessName === "" ||

        //   item.ProcessName.toLowerCase().includes(filters.ProcessName.toLowerCase())) &&

        (filters.Status === "" ||
          item.Status.toLowerCase().includes(filters.Status.toLowerCase())) &&
        (filters.RequestedDate === "" ||
          new Date(item.Created)
            .toLocaleDateString()
            .startsWith(filters.RequestedDate + "")) &&
        (filters.Title === "" ||
          (activeTab == "Automation"
            ? item?.ApprovalTitle?.toLowerCase().includes(
              filters.Title.toLowerCase()
            )
            : item?.Title?.toLowerCase().includes(
              filters.Title.toLowerCase()
            ))) &&
        (filters.RequestedBy === "" ||
          (activeTab == "Automation"
            ? item?.Author?.Title?.toLowerCase().includes(
              filters.RequestedBy.toLowerCase()
            )
            : item?.Requester?.Title?.toLowerCase().includes(
              filters.RequestedBy.toLowerCase()
            )))
      );
    });

    const sortedData = filteredData?.sort((a, b) => {
      if (sortConfig.key === "SNo") {
        // Sort by index

        const aIndex = data.indexOf(a);

        const bIndex = data.indexOf(b);

        return sortConfig.direction === "ascending"
          ? aIndex - bIndex
          : bIndex - aIndex;
      } else if (sortConfig.key == "RequestedDate") {
        // Sort by other keys

        const aValue = a["Created"] ? new Date(a["Created"]) : "";

        const bValue = b["Created"] ? new Date(b["Created"]) : "";

        if (aValue < bValue) {
          return sortConfig.direction === "ascending" ? -1 : 1;
        }

        if (aValue > bValue) {
          return sortConfig.direction === "ascending" ? 1 : -1;
        }
      } else if (sortConfig.key) {
        // Sort by other keys

        const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : "";

        const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : "";

        if (aValue < bValue) {
          return sortConfig.direction === "ascending" ? -1 : 1;
        }

        if (aValue > bValue) {
          return sortConfig.direction === "ascending" ? 1 : -1;
        }
      }

      return 0;
    });

    return sortedData;
  };

  const filteredMyApprovalData = applyFiltersAndSorting(myApprovalsData);

  const filteredNewsData = applyFiltersAndSorting(newsData);

  const [currentPage, setCurrentPage] = React.useState(1);
  const [currentGroup, setCurrentGroup] = React.useState(1);
  const itemsPerPage = 10;
  const pagesPerGroup = 10;
  const totalPages = Math.ceil(filteredMyApprovalData?.length / itemsPerPage);
  const totalGroups = Math.ceil(totalPages / pagesPerGroup);
  const [ContentData, setContentData] = React.useState<any>([]);

  const [currentItem, setCurrentItem] = React.useState<any>([]);

  // const handlePageChange = (pageNumber: any) => {
  //   if (pageNumber > 0 && pageNumber <= totalPages) {
  //     setCurrentPage(pageNumber);
  //   }
  // };
  const handleGroupChange = (direction: "next" | "prev") => {
    const newGroup = currentGroup + (direction === "next" ? 1 : -1);
    if (newGroup > 0 && newGroup <= totalGroups) {
      setCurrentGroup(newGroup);
      setCurrentPage((newGroup - 1) * pagesPerGroup + 1); // Go to the first page of the new group
    }
  };

  const handlePageChange = (pageNumber: any) => {

    if (pageNumber > 0 && pageNumber <= totalPages) {

      setCurrentPage(pageNumber);
      const newGroup = Math.ceil(pageNumber / pagesPerGroup);
      setCurrentGroup(newGroup);
    }

  };
  const startIndex = (currentPage - 1) * itemsPerPage;

  const endIndex = startIndex + itemsPerPage;

  const currentData = filteredMyApprovalData?.slice(startIndex, endIndex);
  const startPage = (currentGroup - 1) * pagesPerGroup + 1;
  const endPage = Math.min(currentGroup * pagesPerGroup, totalPages);
  const newsCurrentData = filteredNewsData?.slice(startIndex, endIndex);

  const [editID, setEditID] = React.useState(null);

  const [ImagepostIdsArr, setImagepostIdsArr] = React.useState([]);

  const siteUrl = props.siteUrl;

  const Breadcrumb = [
    {
      MainComponent: "Home",

      MainComponentURl: `${siteUrl}/SitePages/Dashboard.aspx`,
    },

    {
      ChildComponent: "My Approvals",

      ChildComponentURl: `${siteUrl}/SitePages/MyApprovals.aspx`,
    },
  ];

  console.log(announcementData, "announcementData");

  const exportToExcel = (data: any[], fileName: string) => {
    const workbook = XLSX.utils.book_new();

    const worksheet = XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  };
  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(1);
      setCurrentGroup(1);
    }
  }, [filteredMyApprovalData, currentGroup]);
  // const fetchOptions = async () => {

  //   try {

  //     const items = await fetchUserInformationList(sp);

  //     console.log(items, "itemsitemsitems");

  //     const formattedOptions = items.map((item: { Title: any; Id: any }) => ({

  //       name: item.Title, // Adjust according to your list schema

  //       id: item.Id,

  //     }));

  //     setOpions(formattedOptions);

  //   } catch (error) {

  //     console.error("Error fetching options:", error);

  //   }

  // };

  const handleSortChange = (key: string) => {
    let direction = "ascending";

    if (sortConfig.key === key && sortConfig.direction === "ascending") {
      direction = "descending";
    }

    setSortConfig({ key, direction });
  };

  React.useEffect(() => {
    const showNavbar = (
      toggleId: string,

      navId: string,

      bodyId: string,

      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);

      const nav = document.getElementById(navId);

      const bodypd = document.getElementById(bodyId);

      const headerpd = document.getElementById(headerId);

      UserGet();

      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener("click", () => {
          nav.classList.toggle("show");

          toggle.classList.toggle("bx-x");

          bodypd.classList.toggle("body-pd");

          headerpd.classList.toggle("body-pd");
        });
      }
    };

    showNavbar("header-toggle", "nav-bar", "body-pd", "header");

    const linkColor = document.querySelectorAll(".nav_link");

    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach((l) => l.classList.remove("active"));

        this.classList.add("active");
      }
    }

    linkColor.forEach((l) => l.addEventListener("click", colorLink));
  }, [useHide]);

  const UserGet = async () => {
    const users = await sp.web.siteUsers();

    console.log(users, "users");
  };

  const handleRedirect = async (
    e: React.MouseEvent<SVGElement, MouseEvent>,
    Item: any, mode: any
  ) => {
    e.preventDefault();

    let arr = [];

    setCurrentItem(Item);



    //setisActivedata(true);

    let sessionkey = "";
    let redirecturl = "";
    if (activeTab == "Automation") {
      window.open(Item.RedirectionLink, "_blank");
      //window.location.href = `${Item.RedirectionLink}`;
    } else if (activeTab == "Intranet") {
      setContentData(await getDataByID(sp, Item?.ContentId, Item?.ContentName));

      if (Item?.ProcessName !== "Blog") {
        setisActivedata(true);
      }
      if (Item?.ProcessName) {
        switch (Item?.ProcessName) {
          case "Announcement":
            sessionkey = "announcementId";
            redirecturl =
              `${siteUrl}/SitePages/AddAnnouncement.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "News":
            sessionkey = "announcementId";
            redirecturl =
              `${siteUrl}/SitePages/AddAnnouncement.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Event":
            sessionkey = "EventId";
            redirecturl =
              `${siteUrl}/SitePages/EventMasterForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Media":
            sessionkey = "mediaId";
            redirecturl =
              `${siteUrl}/SitePages/MediaGalleryForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Blog":
            sessionkey = "blogId";
            redirecturl =
              `${siteUrl}/SitePages/BlogDetails.aspx?` + Item?.ContentId + "&page=MyApproval";
            break;
          default:
        }

        const encryptedId = encryptId(String(Item?.ContentId));
        sessionStorage.setItem(sessionkey, encryptedId);
        location.href = redirecturl;
      }
    }

    // const encryptedId = encryptId(String(Item?.ContentId));

    // sessionStorage.setItem("announcementId", encryptedId);
  };

  const handleFromSubmit = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    Status: string
  ) => {
    e.preventDefault();

    const postPayload = {
      Remark: formData.Remark,

      Status: Status,
      TriggerUpdateFlow: true,
    };

    console.log(postPayload);

    const postResult = await updateItemApproval(
      postPayload,
      sp,
      currentItem.Id
    );

    if (postResult) {
      setTimeout(() => {
        window.location.reload();
      }, 1000);
    }
  };

  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(1);
    }
  }, [filteredMyApprovalData]);
  return (
    <div id="wrapper" ref={elementRef}>
      <div className="app-menu" id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>

      <div className="content-page">
        <HorizontalNavbar _context={sp} siteUrl={siteUrl} />

        <div
          className="content"
          style={{
            marginLeft: `${!useHide ? "240px" : "80px"}`,

            marginTop: "0rem",
          }}
        >
          <div className="container-fluid paddb">
            <div className="row" style={{ paddingLeft: "0.5rem" }}>
              <div className="col-md-4">

                <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp} />

              </div>

              <div className="col-md-8">
                <div className="row">
                  <div style={{ textAlign: "center" }} className="col-md-3 newtexleft">
                    <div className="mb-0">
                      <label htmlFor="Status" className="form-label mt-2">
                        <img src={require('../assets/delegation.png')}
                          className='me-1' alt="d" />   Acting on behalf of
                      </label>
                    </div>
                  </div>
                  <div style={{ paddingLeft: '0px' }} className="col-md-4">
                    <select
                      id="Type"
                      name="Type"
                      onChange={(e) => handleStatusChange(e.target.name, 'Pending', e.target.value)}
                      className="form-select"
                      disabled={loading || actingForUser.length === 0}
                    >
                      <option value="">
                        {loading ? "Loading..." : actingForUser.length === 0 ? "No Delegation" : "Select an option"}
                      </option>
                      {actingForUser.map((item, index) => (
                        <option key={index} value={item.email}>
                          {item.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ textAlign: "center", padding: '0px' }} className="col-md-1 newtexleft ivon">
                    <div className="mb-0">
                      <label htmlFor="Status" className="form-label newfil mt-0 mb-0">
                        <svg fill="#3c3c3c" width="23px" height="36px" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg" stroke="#b3b3b3"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path d="M12,25l6.67,6.67a1,1,0,0,0,.7.29.91.91,0,0,0,.39-.08,1,1,0,0,0,.61-.92V13.08L31.71,1.71A1,1,0,0,0,31.92.62,1,1,0,0,0,31,0H1A1,1,0,0,0,.08.62,1,1,0,0,0,.29,1.71L11.67,13.08V24.33A1,1,0,0,0,12,25ZM3.41,2H28.59l-10,10a1,1,0,0,0-.3.71V28.59l-4.66-4.67V12.67a1,1,0,0,0-.3-.71Z"></path> </g></svg>
                      </label>
                    </div>
                  </div>
                  <div style={{ paddingLeft: '0px' }} className="col-md-4 ivon2">
                    <select
                      id="Type"
                      name="Type"
                      onChange={(e) =>
                        handleStatusChange(e.target.name, e.target.value, actingforuseremail)
                      }
                      className="form-select"
                    >
                      {/* <option value="">Pending</option> */}
                      {StatusTypeData.map((item, index) => (
                        <option key={index} value={item.name}>
                          {item.name}
                        </option>
                      ))}
                    </select>
                  </div>

                </div>


              </div>
            </div>

            <div className="row mt-4">
              <div className="col-12">
                <div className="card mb-0 cardcsss">
                  <div className="card-body">
                    <div className="d-flex flex-wrap align-items-center justify-content-center">
                      <ul
                        className="nav nav-pills navtab-bg float-end justify-content-center"
                        role="tablist"
                      >
                        <li className="nav-item myapprovalcomingsoon" role="presentation">
                          {/* <li className="nav-item" role="presentation"> */}
                          <a
                            //onClick={() => handleTabClick("Intranet")}
                            className={`nav-link ${activeTab === "Intranet" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Intranet"}
                            role="tab"
                          >
                            <span className="lenbg1">Intranet</span>{" "}
                            <span className="lenbg comingsoone">
                              {" "}
                              {/* {myApprovalsDataAll.length}  */}
                              coming soon
                            </span>
                          </a>
                          {/* <a
                            onClick={() => handleTabClick("Intranet")}
                            className={`nav-link ${activeTab === "Intranet" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Intranet"}
                            role="tab"
                          >
                            <span className="lenbg1">Intranet</span>{" "}
                            <span className="lenbg">
                              {" "}
                              {myApprovalsDataAll.length}
                            </span>
                          </a> */}
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("DMS")}
                            className={`nav-link ${activeTab === "DMS" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "DMS"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">DMS </span>
                            <span className="lenbg">
                              {Mylistdata.length}
                            </span>
                          </a>
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("Automation")}
                            className={`nav-link ${activeTab === "Automation" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Automation"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">Automation </span>
                            <span className="lenbg">
                              {myApprovalsDataAutomation.length}
                            </span>
                          </a>
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("ProjectWorkflow")}
                            className={`nav-link ${activeTab === "ProjectWorkflow" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "ProjectWorkflow"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">Project Workflow </span>
                            <span className="lenbg">
                              {ProjectWorkflowdata.length}
                            </span>
                          </a>
                        </li>
                      </ul>
                    </div>
                  </div>
                </div>
                <div></div>
              </div>
            </div>
            {(activeTab === "Intranet" ||
              activeTab === "Automation" ||
              activeTab === "DMS" ||
              activeTab === "ProjectWorkflow") && (
                <div>
                  {!isActivedata && (
                    <div className=" mt-2">
                      <div className="">
                        <div id="cardCollpase4" className="collapse show">
                          <div className="table-responsive pt-0">
                            {activeTab === "Intranet" ||
                              activeTab === "Automation" ? (

                              <>
                                <table
                                  className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                  style={{ position: "relative" }}
                                >
                                  <thead>
                                    <tr>
                                      <th
                                        style={{
                                          borderBottomLeftRadius: "0px",

                                          minWidth: "40px",

                                          maxWidth: "40px",

                                          borderTopLeftRadius: "0px",
                                        }}
                                      >
                                        <div
                                          className="d-flex pb-2"
                                          style={{ justifyContent: "space-evenly" }}
                                        >
                                          <span>S.No.</span>

                                          <span
                                            onClick={() => handleSortChange("SNo")}
                                          >
                                            <FontAwesomeIcon icon={faSort} />
                                          </span>
                                        </div>

                                        <div className="bd-highlight">
                                          <input
                                            type="text"
                                            placeholder="index"
                                            onChange={(e) =>
                                              handleFilterChange(e, "SNo")
                                            }
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter' && !e.shiftKey) {
                                                e.preventDefault(); // Prevents the new line in textarea
                                              }
                                            }}
                                            className="inputcss"
                                            style={{ width: "100%" }}
                                          />
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "80px",
                                          maxWidth: "80px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Request ID</span>

                                            <span
                                              onClick={() =>
                                                handleSortChange("RequestID")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Request ID"
                                              onChange={(e) =>
                                                handleFilterChange(e, "RequestID")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>
                                      {/* {activeTab == "Intranet" && ( */}
                                      <th
                                        style={{
                                          minWidth: "120px",
                                          maxWidth: "120px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Title</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("Title")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Title"
                                              onChange={(e) =>
                                                handleFilterChange(e, "Title")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>
                                      {/* )} */}
                                      <th
                                        style={{
                                          minWidth: "120px",
                                          maxWidth: "120px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Process Name</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("ProcessName")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Process Name"
                                              onChange={(e) =>
                                                handleFilterChange(e, "ProcessName")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "100px",
                                          maxWidth: "100px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Requested By</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("RequestedBy")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested By"
                                              onChange={(e) =>
                                                handleFilterChange(e, "RequestedBy")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "130px",
                                          maxWidth: "130px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Requested Date</span>{" "}
                                            {/* <span
                                            onClick={() =>
                                              handleSortChange("RequestedDate")
                                            }
                                          >
                                            <FontAwesomeIcon icon={faSort} />{" "}
                                          </span> */}
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested Date"
                                              onChange={(e) =>
                                                handleFilterChange(
                                                  e,
                                                  "RequestedDate"
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "80px",
                                          maxWidth: "80px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Status</span>{" "}
                                            {/* <span
                                            onClick={() =>
                                              handleSortChange("Status")
                                            }
                                          >
                                            <FontAwesomeIcon icon={faSort} />{" "}
                                          </span> */}
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Status"
                                              onChange={(e) =>
                                                handleFilterChange(e, "Status")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "50px",

                                          maxWidth: "50px",

                                          borderBottomRightRadius: "0px",

                                          borderTopRightRadius: "0px",

                                          textAlign: "center",

                                          verticalAlign: "top",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Action</span>{" "}
                                          </div>
                                        </div>
                                      </th>
                                    </tr>
                                  </thead>
                                  {console.log(
                                    "currentData",
                                    currentData,
                                    isActivedata
                                  )}
                                  <tbody>
                                    {((loading && currentData?.length == 0)
                                      ||
                                      (StatusChange)) && (
                                        <div className="loadernewadd">
                                          <div>
                                            <img
                                              src={require("../../../CustomAsset/birdloader.gif")}
                                              className="alignrightl"
                                              alt="Loading..."
                                            />
                                          </div>
                                          <div className="loadnewarg">
                                            <span>Loading </span>{" "}
                                            <span>
                                              <img
                                                src={require("../../corporateDirectory/assets/argloader.gif")}
                                                className="alignrightbird"
                                                alt="Loading..."
                                              />
                                            </span>
                                          </div>
                                        </div>
                                      )}
                                    {!loading && currentData?.length === 0 ? (
                                      <div

                                        className="no-results card card-body align-items-center  annusvg text-center "

                                        style={{

                                          display: "flex",

                                          justifyContent: "center",
                                          position: 'relative',
                                          marginTop: '10px',
                                          height: '500px'

                                        }}

                                      >
                                        <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>

                                        <p className="font-14 text-muted text-center">No Approval found </p>

                                      </div>
                                    ) : (
                                      !StatusChange && currentData?.map(
                                        (item: any, index: number) => (
                                          <tr
                                            key={index}

                                          >
                                            <td
                                              style={{
                                                minWidth: "40px",
                                                maxWidth: "40px",
                                              }}
                                            >
                                              <div
                                                style={{ marginLeft: "0px" }}
                                                className="indexdesign"
                                              >
                                                {" "}
                                                {startIndex + index + 1}
                                              </div>{" "}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",

                                                maxWidth: "80px",
                                                textAlign: 'center',

                                                textTransform: "capitalize",
                                              }}
                                              title={item.RequestID}
                                            >
                                              {item.RequestID}
                                            </td>
                                            {/* {activeTab == "Intranet" && ( */}
                                            <td
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                              }}
                                              title={activeTab == "Intranet" ? item.Title : item.ApprovalTitle}
                                            >
                                              {activeTab == "Intranet" ? item.Title : item.ApprovalTitle}
                                            </td>
                                            {/* )} */}
                                            <td
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <span className="badge font-12 bg-secondary">   {item.ProcessName}</span>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "100px",
                                                maxWidth: "100px",
                                              }}
                                              title={activeTab == "Automation"
                                                ? item?.Author?.Title
                                                : item?.Requester?.Title}
                                            >
                                              {activeTab == "Automation"
                                                ? item?.Author?.Title
                                                : item?.Requester?.Title}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "130px",
                                                maxWidth: "130px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <div className="btn btn-light1">
                                                {/* {new Date(
                                                  item?.Created
                                                ).toLocaleDateString()} */}
                                                {new Date(item?.Created).toLocaleString('en-US', {
                                                  month: '2-digit',
                                                  day: '2-digit',
                                                  year: 'numeric',
                                                  hour: '2-digit',
                                                  minute: '2-digit',
                                                  //second: '2-digit',
                                                  hour12: true
                                                })}
                                              </div>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <div className="btn btn-status">
                                                {item?.Status}
                                              </div>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "50px",
                                                maxWidth: "50px",
                                                textAlign: 'center'
                                              }}
                                              className="fe-eye font-18"
                                            >
                                              {item?.Status.toLowerCase() == "approved" || item?.Status.toLowerCase() == "rejected"
                                                || item?.Status.toLowerCase() == "completed" ?


                                                <Eye onClick={(e) =>
                                                  handleRedirect(e, item, "view")
                                                }

                                                  style={{

                                                    minWidth: "20px",

                                                    maxWidth: "20px",



                                                    cursor: "pointer",

                                                  }} />

                                                :
                                                <Edit
                                                  onClick={(e) =>
                                                    handleRedirect(e, item, "approval")
                                                  }

                                                  style={{
                                                    marginLeft: "0px",

                                                    cursor: "pointer",
                                                  }}
                                                />
                                              }
                                            </td>
                                          </tr>
                                        )
                                      )
                                    )}
                                  </tbody>
                                </table>

                                {currentData?.length > 0 ? (
                                  <nav className="pagination-container">
                                    <ul className="pagination">
                                      {/* <li
        className={`page-item ${currentPage === 1 ? "disabled" : ""
          }`}
      > */}
                                      <li

                                        className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""
                                          }`}
                                        onClick={() => handleGroupChange("prev")}
                                      >

                                        <a
                                          className="page-link"
                                          // onClick={() =>
                                          //   handlePageChange(currentPage - 1)
                                          // }
                                          aria-label="Previous"
                                        >
                                          «
                                        </a>
                                      </li>
                                      {Array.from(

                                        { length: endPage - startPage + 1 },

                                        (_, num) => {
                                          const pageNum = startPage + num;
                                          return (

                                            <li

                                              key={pageNum}

                                              className={`page-item ${currentPage === pageNum ? "active" : ""

                                                }`}

                                            >

                                              <a

                                                className="page-link"

                                                onClick={() =>

                                                  handlePageChange(pageNum)

                                                }

                                              >

                                                {pageNum}

                                              </a>

                                            </li>

                                          )
                                        }

                                      )}

                                      <li

                                        className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""

                                          }`}
                                        onClick={() => handleGroupChange("next")}
                                      >

                                        <a

                                          className="page-link"

                                          onClick={() =>

                                            handlePageChange(currentPage + 1)

                                          }

                                          aria-label="Next"

                                        >

                                          »

                                        </a>

                                      </li>

                                    </ul>
                                  </nav>
                                ) : (
                                  <></>
                                )}
                              </>
                            ) : null}

                            {activeTab === "DMS" && (
                              <div>
                                {!showNestedDMSTable ? (
                                  <div>
                                    <table
                                      className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                      style={{ position: "relative" }}
                                    >
                                      <thead>
                                        <tr>
                                          <th
                                            style={{
                                              borderBottomLeftRadius: "0px",

                                              minWidth: "40px",

                                              maxWidth: "40px",

                                              borderTopLeftRadius: "0px",
                                            }}
                                          >
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>S.No.</span>

                                              <span
                                                onClick={() =>
                                                  handleSortChange("SNo")
                                                }
                                              >
                                                <FontAwesomeIcon icon={faSort} />
                                              </span>
                                            </div>

                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                placeholder="index"
                                                onChange={(e) =>
                                                  handleFilterChange(e, "SNo")
                                                }
                                                onKeyDown={(e) => {
                                                  if (e.key === 'Enter' && !e.shiftKey) {
                                                    e.preventDefault();
                                                  }
                                                }}
                                                className="inputcss"
                                                style={{ width: "100%" }}
                                              />
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "80px",
                                              maxWidth: "80px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Request ID</span>

                                                <span
                                                  onClick={() =>
                                                    handleSortChange("RequestID")
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />
                                                </span>
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Request ID"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "RequestID"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>
                                          <th
                                            style={{
                                              minWidth: "120px",
                                              maxWidth: "120px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Title</span>

                                                <span
                                                  onClick={() =>
                                                    handleSortChange("Title")
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />
                                                </span>
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Title"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "RequestID"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "120px",
                                              maxWidth: "120px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex  pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Process Name</span>{" "}
                                                <span
                                                  onClick={() =>
                                                    handleSortChange(
                                                      "ProcessName"
                                                    )
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />{" "}
                                                </span>
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Process Name"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "ProcessName"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "100px",
                                              maxWidth: "100px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex  pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Requested By</span>{" "}
                                                <span
                                                  onClick={() =>
                                                    handleSortChange(
                                                      "RequestedBy"
                                                    )
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />{" "}
                                                </span>
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Requested By"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "RequestedBy"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "130px",
                                              maxWidth: "130px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex  pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Requested Date</span>{" "}
                                                {/* <span
                                                  onClick={() =>
                                                    handleSortChange(
                                                      "RequestedDate"
                                                    )
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />{" "}
                                                </span> */}
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Requested Date"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "RequestedDate"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "80px",
                                              maxWidth: "80px",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex  pb-2"
                                                style={{
                                                  justifyContent: "space-evenly",
                                                }}
                                              >
                                                <span>Status</span>{" "}
                                                {/* <span
                                                  onClick={() =>
                                                    handleSortChange("Status")
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />{" "}
                                                </span> */}
                                              </div>

                                              <div className=" bd-highlight">
                                                <input
                                                  type="text"
                                                  placeholder="Filter by Status"
                                                  onChange={(e) =>
                                                    handleFilterChange(
                                                      e,
                                                      "Status"
                                                    )
                                                  }
                                                  onKeyDown={(e) => {
                                                    if (e.key === 'Enter' && !e.shiftKey) {
                                                      e.preventDefault();
                                                    }
                                                  }}
                                                  className="inputcss"
                                                  style={{ width: "100%" }}
                                                />
                                              </div>
                                            </div>
                                          </th>

                                          <th
                                            style={{
                                              minWidth: "50px",

                                              maxWidth: "50px",

                                              borderBottomRightRadius: "0px",

                                              borderTopRightRadius: "0px",

                                              textAlign: "center",

                                              verticalAlign: "top",
                                            }}
                                          >
                                            <div className="d-flex flex-column bd-highlight ">
                                              <div
                                                className="d-flex  pb-2"
                                                style={{
                                                  justifyContent: "space-between",
                                                }}
                                              >
                                                <span>Action</span>{" "}
                                              </div>
                                            </div>
                                          </th>
                                        </tr>
                                      </thead>
                                      {console.log(
                                        "currentData",
                                        currentData,
                                        isActivedata
                                      )}
                                      <tbody>
                                        {((loading && currentData?.length == 0) ||
                                          (StatusChange)) && (
                                            <div className="loadernewadd">
                                              <div>
                                                <img
                                                  src={require("../../../CustomAsset/birdloader.gif")}
                                                  className="alignrightl"
                                                  alt="Loading..."
                                                />
                                              </div>
                                              <div className="loadnewarg">
                                                <span>Loading </span>{" "}
                                                <span>
                                                  <img
                                                    src={require("../../corporateDirectory/assets/argloader.gif")}
                                                    className="alignrightbird"
                                                    alt="Loading..."
                                                  />
                                                </span>
                                              </div>
                                            </div>
                                          )}
                                        {!loading && currentData?.length === 0 ? (
                                          <div

                                            className="no-results card card-body align-items-center  annusvg text-center "

                                            style={{

                                              display: "flex",

                                              justifyContent: "center",
                                              position: 'relative',
                                              marginTop: '10px',
                                              height: '500px'

                                            }}

                                          >
                                            <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>

                                            <p className="font-14 text-muted text-center">No Approval found </p>

                                          </div>
                                        ) : (
                                          !StatusChange && currentData?.map(
                                            (item: any, index: number) => (
                                              <tr
                                                key={index}

                                              >
                                                <td
                                                  style={{
                                                    minWidth: "40px",
                                                    maxWidth: "40px",

                                                    backgroundColor: "transparent",

                                                  }}
                                                >
                                                  <div
                                                    style={{ marginLeft: "0px" }}
                                                    className="indexdesign"
                                                  >
                                                    {" "}
                                                    {startIndex + index + 1}
                                                  </div>{" "}
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "80px",

                                                    maxWidth: "80px",
                                                    textAlign: 'center',

                                                    textTransform: "capitalize",
                                                  }}
                                                  title={item?.FileUID?.RequestNo}
                                                >
                                                  {item?.FileUID?.RequestNo}

                                                </td>
                                                <td
                                                  style={{
                                                    minWidth: "120px",

                                                    maxWidth: "120px",


                                                  }}
                                                  title={item?.FileUID?.FileName}
                                                >
                                                  {item?.FileUID?.FileName}
                                                </td>
                                                <td
                                                  style={{
                                                    minWidth: "120px",
                                                    maxWidth: "120px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                  <span className="badge font-12 bg-secondary">  {item?.FileUID?.Processname} </span>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "100px",
                                                    maxWidth: "100px",
                                                  }}
                                                >
                                                  {item?.RequestedByTitle}
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "130px", maxWidth: "130px",
                                                    // maxWidth: "100px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                  <div className="btn btn-light1">

                                                    {/* {item?.FileUID?.Created} */}
                                                    {new Date(item?.FileUID?.Created).toLocaleString('en-US', {
                                                      month: '2-digit',
                                                      day: '2-digit',
                                                      year: 'numeric',
                                                      hour: '2-digit',
                                                      minute: '2-digit',
                                                      //second: '2-digit',
                                                      hour12: true
                                                    })}
                                                  </div>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "80px",
                                                    maxWidth: "80px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                  <div className="btn btn-status">
                                                    {item?.FileUID?.Status}
                                                  </div>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "50px",
                                                    maxWidth: "50px",
                                                  }}
                                                  className="fe-eye font-18"
                                                >

                                                  <Edit
                                                    onClick={(e) => { getTaskItemsbyID(e, item?.FileUID?.FileUID, item?.FileUID?.Processname); handleShowNestedDMSTable() }}
                                                    style={{


                                                      marginLeft: "15px",

                                                      cursor: "pointer",
                                                    }}
                                                  />
                                                </td>
                                              </tr>
                                            )
                                          )
                                        )}
                                      </tbody>
                                    </table>
                                    {currentData?.length > 0 ? (
                                      <nav className="pagination-container">
                                        <ul className="pagination">
                                          {/* <li
        className={`page-item ${currentPage === 1 ? "disabled" : ""
          }`}
      > */}
                                          <li

                                            className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""
                                              }`}
                                            onClick={() => handleGroupChange("prev")}
                                          >

                                            <a
                                              className="page-link"
                                              // onClick={() =>
                                              //   handlePageChange(currentPage - 1)
                                              // }
                                              aria-label="Previous"
                                            >
                                              «
                                            </a>
                                          </li>
                                          {Array.from(

                                            { length: endPage - startPage + 1 },

                                            (_, num) => {
                                              const pageNum = startPage + num;
                                              return (

                                                <li

                                                  key={pageNum}

                                                  className={`page-item ${currentPage === pageNum ? "active" : ""

                                                    }`}

                                                >

                                                  <a

                                                    className="page-link"

                                                    onClick={() =>

                                                      handlePageChange(pageNum)

                                                    }

                                                  >

                                                    {pageNum}

                                                  </a>

                                                </li>

                                              )
                                            }

                                          )}

                                          <li

                                            className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""

                                              }`}
                                            onClick={() => handleGroupChange("next")}
                                          >

                                            <a

                                              className="page-link"

                                              onClick={() =>

                                                handlePageChange(currentPage + 1)

                                              }

                                              aria-label="Next"

                                            >

                                              »

                                            </a>

                                          </li>

                                        </ul>
                                      </nav>
                                    ) : (
                                      <></>
                                    )}

                                  </div>
                                ) : (
                                  <div>
                                    {folderActionOrFileAction === "New File Request" && (
                                      <>
                                        <DMSMyApprovalAction props={{ currentItemID, actingforuseremail }} />
                                        <div className="col-sm-12 text-center">
                                          {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                          <button type="button" className="btn cancel-btn newp waves-effect waves-light m-3" style={{ fontSize: '0.875rem' }} onClick={() => setShowNestedDMSTable(false)}>
                                            <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                              className='me-1' alt="x" />
                                            Cancel
                                          </button>
                                        </div>
                                      </>
                                    )}
                                    {folderActionOrFileAction === "New Folder Request" && (
                                      <>
                                        <DMSMyFolderApprovalAction props={{ currentItemID, actingforuseremail }} />
                                        <div className="col-sm-12 text-center">
                                          {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                          <button type="button" className="btn cancel-btn newp waves-effect waves-light m-3" style={{ fontSize: '0.875rem' }} onClick={() => setShowNestedDMSTable(false)}>
                                            <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                              className='me-1' alt="x" />
                                            Cancel
                                          </button>
                                        </div>
                                      </>
                                    )}

                                  </div>
                                )}
                              </div>
                            )}

                            {/* {activeTab === "ProjectWorkflow" && (
                              <div>
                                <table className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0">
                                  <thead>
                                    <tr>
                                      <th style={{ borderBottomLeftRadius: "0px", minWidth: "40px", maxWidth: "40px", borderTopLeftRadius: "0px" }}>
                                        <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                          <span>S.No.</span>
                                          <span onClick={() => handleSortChange("SNo")}>
                                            <FontAwesomeIcon icon={faSort} />
                                          </span>
                                        </div>
                                        <div className="bd-highlight">
                                          <input
                                            type="text"
                                            placeholder="index"
                                            onChange={(e) => handleFilterChange(e, "SNo")}
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter' && !e.shiftKey) {
                                                e.preventDefault();
                                              }
                                            }}
                                            className="inputcss"
                                            style={{ width: "100%" }}
                                          />
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Request ID</span>
                                            <span onClick={() => handleSortChange("RequestID")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Request ID"
                                              onChange={(e) => handleFilterChange(e, "RequestID")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Project Name</span>
                                            <span onClick={() => handleSortChange("Title")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Project Name"
                                              onChange={(e) => handleFilterChange(e, "Title")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Process Name</span>
                                            <span onClick={() => handleSortChange("ProcessName")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Process Name"
                                              onChange={(e) => handleFilterChange(e, "ProcessName")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "100px", maxWidth: "100px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Requested By</span>
                                            <span onClick={() => handleSortChange("RequestedBy")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested By"
                                              onChange={(e) => handleFilterChange(e, "RequestedBy")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "130px", maxWidth: "130px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Requested Date</span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested Date"
                                              onChange={(e) => handleFilterChange(e, "RequestedDate")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Status</span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Status"
                                              onChange={(e) => handleFilterChange(e, "Status")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "50px", maxWidth: "50px", borderBottomRightRadius: "0px", borderTopRightRadius: "0px", textAlign: "center", verticalAlign: "top" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Action</span>
                                          </div>
                                        </div>
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {((loading && currentData?.length == 0) || (StatusChange)) && (
                                      <div className="loadernewadd">
                                        <div>
                                          <img
                                            src={require("../../../CustomAsset/birdloader.gif")}
                                            className="alignrightl"
                                            alt="Loading..."
                                          />
                                        </div>
                                        <div className="loadnewarg">
                                          <span>Loading </span>{" "}
                                          <span>
                                            <img
                                              src={require("../../corporateDirectory/assets/argloader.gif")}
                                              className="alignrightbird"
                                              alt="Loading..."
                                            />
                                          </span>
                                        </div>
                                      </div>
                                    )}
                                    {!loading && currentData?.length === 0 ? (
                                      <div className="no-results card card-body align-items-center annusvg text-center" style={{ display: "flex", justifyContent: "center", position: 'relative', marginTop: '10px', height: '500px' }}>
                                        <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>
                                        <p className="font-14 text-muted text-center">No Project Approvals found</p>
                                      </div>
                                    ) : (
                                      !StatusChange && currentData?.map((item: any, index: number) => (
                                        <tr key={index}>
                                          <td style={{ minWidth: "40px", maxWidth: "40px" }}>
                                            <div style={{ marginLeft: "0px" }} className="indexdesign">
                                              {startIndex + index + 1}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center', textTransform: "capitalize" }} title={item.RequestID}>
                                            {item.RequestID}
                                          </td>
                                          <td style={{ minWidth: "120px", maxWidth: "120px" }} title={item.ProjectName}>
                                            {item.ProjectName}
                                          </td>
                                          <td style={{ minWidth: "120px", maxWidth: "120px", textAlign: 'center' }}>
                                            <span className="badge font-12 bg-secondary">{item.ProcessName}</span>
                                          </td>
                                          <td style={{ minWidth: "100px", maxWidth: "100px" }} title={item.Requester?.Title}>
                                            {item.Requester?.Title}
                                          </td>
                                          <td style={{ minWidth: "130px", maxWidth: "130px", textAlign: 'center' }}>
                                            <div className="btn btn-light1">
                                              {new Date(item.Created).toLocaleString('en-US', {
                                                month: '2-digit',
                                                day: '2-digit',
                                                year: 'numeric',
                                                hour: '2-digit',
                                                minute: '2-digit',
                                                hour12: true
                                              })}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center' }}>
                                            <div className="btn btn-status">
                                              {item.Status}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "50px", maxWidth: "50px", textAlign: 'center' }} className="fe-eye font-18">
                                            {item.Status.toLowerCase() == "approved" || item.Status.toLowerCase() == "rejected" || item.Status.toLowerCase() == "completed" ?
                                              <Eye onClick={(e) => handleProjectWorkflowAction(e, item, "view")} style={{ minWidth: "20px", maxWidth: "20px", cursor: "pointer" }} />
                                              :
                                              <Edit onClick={(e) => handleProjectWorkflowAction(e, item, "approval")} style={{ marginLeft: "0px", cursor: "pointer" }} />
                                            }
                                          </td>
                                        </tr>
                                      ))
                                    )}
                                  </tbody>
                                </table>

                                {currentData?.length > 0 ? (
                                  <nav className="pagination-container">
                                    <ul className="pagination">
                                      <li className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""}`} onClick={() => handleGroupChange("prev")}>
                                        <a className="page-link" aria-label="Previous">«</a>
                                      </li>
                                      {Array.from({ length: endPage - startPage + 1 }, (_, num) => {
                                        const pageNum = startPage + num;
                                        return (
                                          <li key={pageNum} className={`page-item ${currentPage === pageNum ? "active" : ""}`}>
                                            <a className="page-link" onClick={() => handlePageChange(pageNum)}>{pageNum}</a>
                                          </li>
                                        );
                                      })}
                                      <li className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""}`} onClick={() => handleGroupChange("next")}>
                                        <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">»</a>
                                      </li>
                                    </ul>
                                  </nav>
                                ) : (
                                  <></>
                                )}
                              </div>
                            )} */}

                            {activeTab === "ProjectWorkflow" && (
                              <div>
                                {!showProjectForm ? (
                                  // TABLE VIEW
                                  <div className="card cardCss mt-2">
                                    <div className="card-body">
                                      <div id="cardCollpase4" className="collapse show">
                                        <div className="table-responsive pt-0">
                                          <table className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0" style={{ position: "relative" }}>
                                            <thead>
                                              <tr>
                                                <th style={{ borderBottomLeftRadius: "0px", minWidth: "40px", maxWidth: "40px", borderTopLeftRadius: "0px" }}>
                                                  <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                    <span>S.No.</span>
                                                    <span onClick={() => handleSortChange("SNo")}>
                                                      <FontAwesomeIcon icon={faSort} />
                                                    </span>
                                                  </div>
                                                  <div className="bd-highlight">
                                                    <input
                                                      type="text"
                                                      placeholder="index"
                                                      onChange={(e) => handleFilterChange(e, "SNo")}
                                                      onKeyDown={(e) => {
                                                        if (e.key === 'Enter' && !e.shiftKey) {
                                                          e.preventDefault();
                                                        }
                                                      }}
                                                      className="inputcss"
                                                      style={{ width: "100%" }}
                                                    />
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Request ID</span>
                                                      <span onClick={() => handleSortChange("RequestID")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Request ID"
                                                        onChange={(e) => handleFilterChange(e, "RequestID")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Project Name</span>
                                                      <span onClick={() => handleSortChange("Title")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Project Name"
                                                        onChange={(e) => handleFilterChange(e, "Title")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Process Name</span>
                                                      <span onClick={() => handleSortChange("ProcessName")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Process Name"
                                                        onChange={(e) => handleFilterChange(e, "ProcessName")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "100px", maxWidth: "100px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Requested By</span>
                                                      <span onClick={() => handleSortChange("RequestedBy")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Requested By"
                                                        onChange={(e) => handleFilterChange(e, "RequestedBy")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "130px", maxWidth: "130px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Requested Date</span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Requested Date"
                                                        onChange={(e) => handleFilterChange(e, "RequestedDate")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Status</span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Status"
                                                        onChange={(e) => handleFilterChange(e, "Status")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{
                                                  minWidth: "50px",
                                                  maxWidth: "50px",
                                                  borderBottomRightRadius: "0px",
                                                  borderTopRightRadius: "0px",
                                                  textAlign: "center",
                                                  verticalAlign: "top"
                                                }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Action</span>
                                                    </div>
                                                  </div>
                                                </th>
                                              </tr>
                                            </thead>
                                            <tbody>
                                              {((loading && currentData?.length == 0) || (StatusChange)) && (
                                                <div className="loadernewadd">
                                                  <div>
                                                    <img
                                                      src={require("../../../CustomAsset/birdloader.gif")}
                                                      className="alignrightl"
                                                      alt="Loading..."
                                                    />
                                                  </div>
                                                  <div className="loadnewarg">
                                                    <span>Loading </span>{" "}
                                                    <span>
                                                      <img
                                                        src={require("../../corporateDirectory/assets/argloader.gif")}
                                                        className="alignrightbird"
                                                        alt="Loading..."
                                                      />
                                                    </span>
                                                  </div>
                                                </div>
                                              )}
                                              {!loading && currentData?.length === 0 ? (
                                                <div className="no-results card card-body align-items-center annusvg text-center" style={{
                                                  display: "flex",
                                                  justifyContent: "center",
                                                  position: 'relative',
                                                  marginTop: '10px',
                                                  height: '500px'
                                                }}>
                                                  <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                                    <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
                                                  </svg>
                                                  <p className="font-14 text-muted text-center">No Project Approvals found</p>
                                                </div>
                                              ) : (
                                                !StatusChange && currentData?.map((item: any, index: number) => (
                                                  <tr key={index}>
                                                    <td style={{ minWidth: "40px", maxWidth: "40px" }}>
                                                      <div style={{ marginLeft: "0px" }} className="indexdesign">
                                                        {startIndex + index + 1}
                                                      </div>
                                                    </td>
                                                    <td style={{
                                                      minWidth: "80px",
                                                      maxWidth: "80px",
                                                      textAlign: 'center',
                                                      textTransform: "capitalize"
                                                    }} title={item.RequestID}>
                                                      {item.RequestID}
                                                    </td>
                                                    <td style={{ minWidth: "120px", maxWidth: "120px" }} title={item.ProjectName || item.Title}>
                                                      {item.ProjectName || item.Title}
                                                    </td>
                                                    <td style={{ minWidth: "120px", maxWidth: "120px", textAlign: 'center' }}>
                                                      <span className="badge font-12 bg-secondary">{item.ProcessName}</span>
                                                    </td>
                                                    <td style={{ minWidth: "100px", maxWidth: "100px" }} title={item.Requester?.Title}>
                                                      {item.Requester?.Title}
                                                    </td>
                                                    <td style={{ minWidth: "130px", maxWidth: "130px", textAlign: 'center' }}>
                                                      <div className="btn btn-light1">
                                                        {new Date(item.Created).toLocaleString('en-US', {
                                                          month: '2-digit',
                                                          day: '2-digit',
                                                          year: 'numeric',
                                                          hour: '2-digit',
                                                          minute: '2-digit',
                                                          hour12: true
                                                        })}
                                                      </div>
                                                    </td>
                                                    <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center' }}>
                                                      <div className="btn btn-status">
                                                        {item.Status}
                                                      </div>
                                                    </td>
                                                    <td style={{ minWidth: "50px", maxWidth: "50px", textAlign: 'center' }} className="fe-eye font-18">
                                                      <Edit
                                                        onClick={(e) => handleProjectWorkflowAction(e, item, "view")}
                                                        style={{ cursor: "pointer" }}
                                                      />
                                                    </td>
                                                  </tr>
                                                ))
                                              )}
                                            </tbody>
                                          </table>

                                          {currentData?.length > 0 ? (
                                            <nav className="pagination-container">
                                              <ul className="pagination">
                                                <li className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""}`} onClick={() => handleGroupChange("prev")}>
                                                  <a className="page-link" aria-label="Previous">«</a>
                                                </li>
                                                {Array.from({ length: endPage - startPage + 1 }, (_, num) => {
                                                  const pageNum = startPage + num;
                                                  return (
                                                    <li key={pageNum} className={`page-item ${currentPage === pageNum ? "active" : ""}`}>
                                                      <a className="page-link" onClick={() => handlePageChange(pageNum)}>{pageNum}</a>
                                                    </li>
                                                  );
                                                })}
                                                <li className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""}`} onClick={() => handleGroupChange("next")}>
                                                  <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">»</a>
                                                </li>
                                              </ul>
                                            </nav>
                                          ) : (
                                            <></>
                                          )}
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                ) : (
                                  // FORM VIEW - This is where your form goes
                                  <div className="card mt-2">
                                    <div className="card-body">
                                      <div className="form-header d-flex justify-content-between align-items-center mb-3">
                                        <h4 className="text-dark font-16 fw-bold m-0">Project Approval Details</h4>
                                        {/* <button
                                          className="btn btn-secondary"
                                          onClick={handleProjectBackClick}
                                        >
                                          Back
                                        </button> */}
                                      </div>

                                      <div className="row">
                                        <div className="col-md-4">
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Request ID:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.RequestID || ''}
                                              disabled
                                            />
                                          </div> */}
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Project Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProjectName || selectedProjectTask?.Title || ''}
                                              disabled
                                            />
                                          </div>



                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Process Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProcessName || ''}
                                              disabled
                                            />
                                          </div> */}
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Requested By:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Requester?.Title || ''}
                                              disabled
                                            />
                                          </div> */}
                                        </div>
                                        <div className="col-md-4">
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Requested Date:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={new Date(selectedProjectTask?.Created).toLocaleString() || ''}
                                              disabled
                                            />
                                          </div> */}
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Status:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Status || ''}
                                              disabled
                                            />
                                          </div> */}
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Project Type:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProjectType || 'N/A'}
                                              disabled
                                            />
                                          </div>

                                          </div>
                                          <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Deliverable:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Deliverable || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                      

                                      {/* Additional fields from your data */}
                                    
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Document Type:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.DocType || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                          </div>
                                          <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Document Number:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.DocNumber || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Area:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Area || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                          </div>
                                          <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Organization:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Org || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                               

                                      {/* Client and Prepared By fields */}
                                      
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Client Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ClientName || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Prepared By:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.PreparedBy || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                 

                                      {/* Approval Role and Revision Number */}
                                     
                                        {/* <div className="col-md-6">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Approval Role:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ApprovalRole || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div> */}
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Revision Number:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.RevisionNumber || '0'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                      </div>

                                      {/* Document section */}
                                      {projectDocumentInfo?.documentUrl && (
                                        <div className="row">
                                          <div className="col-12">
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Document:</strong></label>
                                              <div className="document-container">
                                                {projectDocumentInfo ? (
                                                  <div
                                                    className="document-link-container p-3 border rounded bg-light"
                                                    style={{ cursor: 'pointer' }}
                                                    onClick={handleProjectOpenDocument}
                                                    title="Click to open document"
                                                  >
                                                    <div className="d-flex align-items-center">
                                                      <span className="document-icon me-2" style={{ fontSize: '1.5rem' }}>📄</span>
                                                      <div>
                                                        <div className="document-name fw-bold">
                                                          {projectDocumentInfo.fileName || projectDocumentInfo.fileLeafRef}
                                                        </div>
                                                        <div className="document-hint text-muted small">
                                                          Click to open document in new tab
                                                          {projectDocumentInfo.sharedLink && (
                                                            <span className="ms-2">🔗 Shared Link Available</span>
                                                          )}
                                                        </div>
                                                      </div>
                                                    </div>
                                                  </div>
                                                ) : projectDocumentInfo?.documentUrl ? (
                                                  <div
                                                    className="document-link-container p-3 border rounded bg-light"
                                                    style={{ cursor: 'pointer' }}
                                                    onClick={() => {
                                                      if (projectDocumentInfo?.documentUrl) {
                                                        window.open(projectDocumentInfo?.documentUrl, '_blank', 'noopener,noreferrer');
                                                      }
                                                    }}
                                                    title="Click to open document"
                                                  >
                                                    <div className="d-flex align-items-center">
                                                      <span className="document-icon me-2" style={{ fontSize: '1.5rem' }}>📄</span>
                                                      <div>
                                                        <div className="document-name fw-bold">Document Available</div>
                                                        <div className="document-hint text-muted small">Click to open document</div>
                                                      </div>
                                                    </div>
                                                  </div>
                                                ) : (
                                                  <div className="alert alert-info mb-0">
                                                    <small>No document available for this deliverable</small>
                                                  </div>
                                                )}
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      )}

                                      {/* Remarks field (if available) */}
                                      {selectedProjectTask?.Remarks && (
                                        <div className="row">
                                          <div className="col-12">
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Remarks:</strong></label>
                                              <textarea
                                                className="form-control"
                                                value={selectedProjectTask?.Remarks || ''}
                                                disabled
                                                rows={3}
                                              />
                                            </div>
                                          </div>
                                        </div>
                                      )}

                                      {/* <div className="row mt-4">
                                        <div className="col-12 text-center">
                                          <p className="text-muted">
                                            <em>Approval actions will be implemented in the next phase</em>
                                          </p>
                                        </div>
                                      </div> */}
                                      {/* Conditional Fields based on Approver Role */}
                                      <div className="row">
                                        <div className="col-12">
                                          {selectedProjectTask?.ApprovalRole === "Document Controller" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Do you need further approval?</strong></label>
                                              <select
                                                value={projectNeedsFurtherApproval}
                                                onChange={(e) => setProjectNeedsFurtherApproval(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                              </select>
                                            </div>
                                          )}

                                          {selectedProjectTask?.ApprovalRole === "DCC" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Want to publish in Dossier?</strong></label>
                                              <select
                                                value={projectWantsToPublishInDossier}
                                                onChange={(e) => setProjectWantsToPublishInDossier(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                              </select>
                                            </div>
                                          )}

                                          <div className="mb-3">
                                            <label className="form-label"><strong>Remarks</strong></label>
                                            <textarea
                                              value={projectRemarks}
                                              onChange={(e) => setProjectRemarks(e.target.value)}
                                              placeholder="Enter your remarks here..."
                                              rows={4}
                                              className="form-control"
                                            />
                                          </div>
                                        </div>
                                      </div>

                                      {/* Dynamic Approval Hierarchy Table */}
                                      {selectedProjectTask?.ApprovalRole === "Document Controller" && (
                                        <div className="card">
                                          <div className="card-body">
                                        <div className="approval-projectHierarchy mt-4">
                                          <div className="d-flex justify-content-between align-items-center mb-3">
                                            <h5>Approval Hierarchy</h5>
                                            <button
                                              className="btn btn-primary btn-sm"
                                              onClick={() => addNewProjectApprovalRow()}
                                            >
                                              + Add New Row
                                            </button>
                                          </div>
                                          <div className="table-responsive">
                                            <table className="table table-bordered">
                                              <thead className="table-light">
                                                <tr>
                                                  <th style={{ width: '15%' }}>Level</th>
                                                  <th style={{ width: '25%' }}>Approver Role</th>
                                                  <th style={{ width: '30%' }}>Approver</th>
                                                  <th style={{ width: '20%' }}>Approval Criteria</th>
                                                  <th style={{ width: '10%' }}>Action</th>
                                                </tr>
                                              </thead>
                                              <tbody>
                                                {projectHierarchy.map((row, index) => (
                                                  <tr key={index}>
                                                    <td>
                                                      <input
                                                        type="text"
                                                        className="form-control"
                                                        value={row.level}
                                                        disabled
                                                      />
                                                    </td>
                                                    <td>
                                                      <select
                                                        value={row.approverRole}
                                                        onChange={(e) => updateProjectApprovalRow(index, 'approverRole', e.target.value)}
                                                        className="form-select"
                                                      >
                                                        <option value="" disabled>Select Role</option>
                                                        <option value="Project Coordinator">Project Coordinator</option>
                                                        <option value="Project Team">Project Team</option>
                                                        <option value="Project Manager">Project Manager</option>
                                                      </select>
                                                    </td>
                                                    <td>
                                                      <Select
                                                        isMulti
                                                        options={users}
                                                        value={row.assignedTo ? users.filter(user =>
                                                          row.assignedTo?.some((assigned: any) => {
                                                            const assignedId = assigned.ID ? assigned.ID.toString() : assigned.toString();
                                                            return assignedId === user.value;
                                                          })
                                                        ) : []}
                                                        onChange={(selectedOptions: any) => handleProjectApproverChange(index, selectedOptions)}
                                                        placeholder="Select Approver(s)"
                                                        className="people-picker"
                                                        classNamePrefix="react-select"
                                                        closeMenuOnSelect={false}
                                                        isClearable={false}
                                                      />
                                                    </td>
                                                    <td>
                                                      <select
                                                        value={row.approvalCriteria}
                                                        onChange={(e) => updateProjectApprovalRow(index, 'approvalCriteria', e.target.value)}
                                                        className="form-select"
                                                      >
                                                        <option value="" disabled>Select Criteria</option>
                                                        <option value="Everyone">Everyone</option>
                                                        <option value="Anyone">Anyone</option>
                                                      </select>
                                                    </td>
                                                    <td className="text-center">
                                                      <button
                                                        className="btn btn-outline-danger btn-sm"
                                                        onClick={() => deleteProjectApprovalRow(index)}
                                                        title="Delete Row"
                                                      >
                                                        🗑️
                                                      </button>
                                                    </td>
                                                  </tr>
                                                ))}
                                              </tbody>
                                            </table>
                                          </div>
                                        </div>
                                        </div>
                                        </div>
                                      )}

                                      {/* Approval Action Buttons */}
                                      <div className="row mt-4">
                                        <div className="col-12">
                                          <div className="d-flex justify-content-center gap-3">
                                            {selectedProjectTask?.Status === "Pending" && (
                                              <>
                                                {selectedProjectTask?.ApprovalRole === "DCC" ? (
                                                  <>
                                                    <button
                                                      className="btn btn-success"
                                                      onClick={() => handleProjectApprovalAction("Approved")}
                                                    >
                                                      Submit
                                                    </button>
                                                    <button
                                                      className="btn btn-secondary"
                                                      onClick={handleProjectBackClick}
                                                    >
                                                      Cancel
                                                    </button>
                                                  </>
                                                ) : (
                                                  <>
                                                    <button
                                                      className="btn btn-success"
                                                      onClick={() => handleProjectApprovalAction("Approved")}
                                                    >
                                                      Approve
                                                    </button>
                                                    <button
                                                      className="btn btn-danger"
                                                      onClick={() => handleProjectApprovalAction("Rejected")}
                                                    >
                                                      Reject
                                                    </button>
                                                    <button
                                                      className="btn btn-warning"
                                                      onClick={() => handleProjectApprovalAction("Rework")}
                                                    >
                                                      Rework
                                                    </button>
                                                    <button
                                                      className="btn btn-secondary"
                                                      onClick={handleProjectBackClick}
                                                    >
                                                      Cancel
                                                    </button>
                                                  </>
                                                )}
                                              </>
                                            )}

                                            {(selectedProjectTask?.Status === "Approved" ||
                                              selectedProjectTask?.Status === "Rejected" ||
                                              selectedProjectTask?.Status === "Rework") && (
                                                <button className="btn btn-secondary" onClick={handleProjectBackClick}>
                                                  Back
                                                </button>
                                              )}
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                )}
                              </div>
                            )}
                            {/* {activeTab === "DMS" ?
               (
               <div>
                    <div className="DMSMasterContainer">
                <h4 className="page-title fw-bold mb-1 font-20">My Approvals 1</h4>
                <div className="" style={{ backgroundColor: 'white', border:'1px solid #54ade0', marginTop:'20px', borderRadius:'20px', padding: '15px'}}>
                <table className="mtbalenew">
    <thead>
      <tr>
        <th
          style={{
            minWidth: '40px',
            maxWidth: '40px',
         
          }}
        >
          S.No
        </th>
        <th>Request ID</th>
        <th>Process Name</th>
        <th>Requested By</th>
        <th >Requested Date</th>
        <th style={{ minWidth: '80px', maxWidth: '80px' }}>Status</th>
        <th
          style={{
            minWidth: '70px',
            maxWidth: '70px',
         
          }}
        >
          Action
        </th>
      </tr>
    </thead>
    <tbody style={{ maxHeight: '8007px' }}>
       
      {Mylistdata.length > 0  ? Mylistdata.map((item, index) => {
      return(
        <tr>
<td style={{ minWidth: '40px', maxWidth: '40px'}}>
  <span style={{marginLeft:'5px'}} className="indexdesign">{index}</span>
  </td>
<td >{(truncateText(item.FileUID.FileUID, 22))}</td>
<td >{item?.ProcessName}</td>
<td >{(truncateText(item.FileUID.RequestedBy, 22))}</td> 
<td >
<div
  style={{
    padding: '5px',
    border: '1px solid #efefef',
    background: '#fff', fontSize:'14px',
    borderRadius: '30px',
  
  }}
  className="btn btn-light"
>
 {item.FileUID.Created}
</div>
</td>
<td style={{ minWidth: '80px', maxWidth: '80px', textAlign:'center' }}>
<div className="finish mb-0">Pending</div>
</td>
<td style={{ minWidth: '70px', maxWidth: '70px' }}>
  {item?.ProcessName === "DMS Folder Approval" ?
    (<a onClick={(e )=>getTaskItemsbyID2(e , item.FileUID.FileUID)}>
    <FontAwesomeIcon icon={faEye} />
   </a>
   ) : item?.ProcessName === "" || item?.ProcessName === null || item?.ProcessName === undefined ?    (
      <a onClick={(e )=>getTaskItemsbyID(e , item.FileUID.FileUID)}>
 <FontAwesomeIcon icon={faEye} />
</a>
    ) : null
  }

</td>
</tr>
      )

       })
       :""

}

      
   

     
    </tbody>
  </table>
        </div>
              </div>
               </div>
               ) : (
                <div>
                  {activeComponent === 'Approval Action' ? (
                    <div>
                   <button style={{float:'right'}} type="button" className="btn btn-secondary" onClick={()=>handleReturnToMain('')}> Back </button>
                  <DMSMyApprovalAction props={currentItemID}/>
                    </div>
               
                  ) : activeComponent === 'DMS Folder Approval' ? (
                    <div>
<button style={{float:'right'}} type="button" className="btn btn-secondary" onClick={()=>handleReturnToMain('')}> Back </button>
<DMSMyFolderApprovalAction props={currentItemID}/>
                    </div>
                                       
                  ) :null
                  
                  } 
             
                </div>
               
            
               )
               } */}
                          </div>


                        </div>
                      </div>
                    </div>
                  )}

                  {isActivedata == true &&
                    ContentData.length > 0 &&
                    currentItem != null && (
                      <div className="row mt-4">
                        <div className="col-12">
                          <div className="card">
                            <div className="card-body">
                              <h4 className="header-title mb-0">
                                {ContentData[0].Title}
                              </h4>

                              <p className="sub-header">
                                {currentItem.EntityName}
                              </p>

                              <div className="row">
                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Company / Department:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.EntityName}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Date of Request:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.Created}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Status:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.Status}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-6">
                                  <div className="mb-0">
                                    <label className="form-label text-dark font-14">
                                      Content:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {ContentData[0].Title ||
                                          ContentData[0].EventName}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-6">
                                  <div className="mb-0">
                                    <label className="form-label text-dark font-14">
                                      Overview:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {ContentData[0].Overview}
                                      </span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          {ContentData[0]?.Description != null && (
                            <div className="card">
                              <div className="card-body">
                                <h4 className="header-title mb-0">Description</h4>

                                <p className="sub-header">
                                  <div
                                    dangerouslySetInnerHTML={{
                                      __html: ContentData[0].Description,
                                    }}
                                  ></div>
                                </p>
                              </div>
                            </div>
                          )}

                          {currentItem.Status == "Submitted" && (
                            <div className="card">
                              <div className="card-body">
                                <div className="row">
                                  {currentItem.Status == "Submitted" && (
                                    <div className="col-lg-12">
                                      <div className="mb-0">
                                        <label
                                          htmlFor="example-textarea"
                                          className="form-label text-dark font-14"
                                        >
                                          Remarks:
                                        </label>

                                        <textarea
                                          className="form-control"
                                          id="example-textarea"
                                          rows={5}
                                          name="Remark"
                                          value={formData.Remark}
                                          onChange={(e) =>
                                            onChange(
                                              e.target.name,
                                              e.target.value
                                            )
                                          }
                                        ></textarea>
                                      </div>
                                    </div>
                                  )}
                                </div>

                                {currentItem.Status == "Submitted" && (
                                  <div className="row mt-3">
                                    <div className="col-12 text-center">
                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-success waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Approved")
                                          }
                                        >
                                          <i className="fe-check-circle me-1"></i>{" "}
                                          Approve
                                        </button>
                                      </a>

                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-warning waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Rework")
                                          }
                                        >
                                          <i className="fe-corner-up-left me-1"></i>{" "}
                                          Rework
                                        </button>
                                      </a>

                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-danger waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Reject")
                                          }
                                        >
                                          <i className="fe-x-circle me-1"></i>{" "}
                                          Reject
                                        </button>
                                      </a>

                                      <button
                                        type="button"
                                        className="btn cancel-btn waves-effect waves-light m-1"
                                        onClick={(e) => handleCancel()}
                                      >
                                        <i className="fe-x me-1"></i> Cancel
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                </div>
              )}
          </div>
        </div>
      </div>
    </div>
  );
};

const MyApproval: React.FC<IMyApprovalProps> = (props) => (
  <Provider>
    <MyApprovalContext props={props} />
  </Provider>
);

export default MyApproval;
