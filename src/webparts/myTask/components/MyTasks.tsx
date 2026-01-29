// import * as React from 'react';
// import styles from './MyTasks.module.scss';
// import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";
// import "@pnp/sp/files/web";
// import CustomPopup from '../../myProject/components/CustomPopup';
// import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
// import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
// import UserContext from "../../../GlobalContext/context";
// import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../../../CustomCss/mainCustom.scss";
// import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
// import { SPFI } from "@pnp/sp/presets/all";
// import { getSP } from "../loc/pnpjsConfig";

// interface IMyTaskProps {
//   context: WebPartContext;
//   siteUrl: string;
//   // description: string;
//   // siteUrl: string;
//   // userDisplayName: string;
//   // isDarkTheme: boolean;
//   // environmentMessage: string;
//   // hasTeamsContext: boolean;
// }

// interface Task {
//   projectId: any;
//   sno: number;
//   projectName: string;
//   projectType: string;
//   deliverable: string;
//   deliverableId: number;
//   area: string;
//   docType: string;
//   docNumber: string;
//   revisionnumber: string;
//   assignedTo: string;
//   org: string;
//   status: "Pending" | "Approved" | "In-Progress";
//   clientName: string;
//   creationDate: Date | undefined;
// }

// interface AuditHistoryItem {
//   sno: number;
//   approvalLevel: string;
//   assignedTo: string;
//   assignedToRole: string;
//   requestorName: string;
//   requestedDate: string;
//   actionTakenBy: string;
//   actionTakenOn: string;
//   remark: string;
//   status: string;
// }

// interface DocumentComment {
//   id: number;
//   userName: string;
//   commentDate: string;
//   pageNumber: string;
//   revision: string;
//   comment: string;
// }

// const MyTask = ({ props }: any) => {
//   const elementRef = React.useRef<HTMLDivElement>(null);
//   const sp: SPFI = getSP();
//   const siteUrl = props.siteUrl
//   const { useHide }: any = React.useContext(UserContext);
//   const [tasks, setTasks] = React.useState<Task[]>([]);
//   const [currentFilter, setCurrentFilter] = React.useState<string>("Pending");
//   const [filteredTasks, setFilteredTasks] = React.useState<Task[]>([]);
//   const [selectedTask, setSelectedTask] = React.useState<Task | null>(null);
//   const [showForm, setShowForm] = React.useState<boolean>(false);
//   const [auditHistory, setAuditHistory] = React.useState<AuditHistoryItem[]>([]);
//   const [showNoAuditHistory, setShowNoAuditHistory] = React.useState<boolean>(false);
//   const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
//   const [comment, setComment] = React.useState<string>("");
//   const [documentControllerId, setDocumentControllerId] = React.useState<number | null>(null);
//   const [uploadedFileName, setUploadedFileName] = React.useState<string>("");
//   const [uploadedFileUrl, setUploadedFileUrl] = React.useState<string>("");
//   const [documentComments, setDocumentComments] = React.useState<DocumentComment[]>([]);
//   const [allDocumentComments, setAllDocumentComments] = React.useState<DocumentComment[]>([]);
//   const [versionList, setVersionList] = React.useState<string[]>([]);
//   const [selectedVersion, setSelectedVersion] = React.useState<string>("");
//   const [showDocumentComments, setShowDocumentComments] = React.useState<boolean>(false);
//   const [existingDocument, setExistingDocument] = React.useState<{
//     name: string;
//     url: string;
//   } | null>(null);

//   const [popup, setPopup] = React.useState<{
//     isOpen: boolean;
//     type: 'confirmation' | 'validation' | 'success' | 'error';
//     title: string;
//     message: string;
//     onConfirm?: () => void;
//   }>({
//     isOpen: false,
//     type: 'success',
//     title: '',
//     message: ''
//   });

//   const spContext = spfi().using(SPFx(props.context));

//   const getDocumentController = async (): Promise<number | null> => {
//     try {
//       const items = await spContext.web.lists.getByTitle("ProjectConfiguration").items
//         .select(
//           "*",
//           "DocumentController/Title",
//           "DocumentController/EMail",
//           "DocumentController/ID"
//         )
//         .expand("DocumentController")
//         .orderBy("Created", false)()

//       if (items.length > 0) {
//         const documentControllerId = items[0].DocumentControllerId;
//         console.log("Document Controller ID:", documentControllerId);
//         return documentControllerId;
//       }
//       return null;
//     } catch (error) {
//       console.error("Error fetching Document Controller:", error);
//       return null;
//     }
//   };

//   React.useEffect(() => {
//     const initializeData = async () => {
//       const docControllerId = await getDocumentController();
//       setDocumentControllerId(docControllerId);
//       await fetchTasks();
//     };

//     initializeData();
//   }, []);

//   const fetchTasks = async () => {
//     const currentUserId = (await spContext.web.currentUser()).Id;
//     try {
//       const items = await spContext.web.lists.getByTitle("DeliverablesDetails").items.select(
//         "*",
//         "ProjectCreationListID/ID",
//         "ProjectCreationListID/ProjectName",
//         "AssignedTo/Id",
//         "AssignedTo/Title"
//       ).expand(
//         "ProjectCreationListID",
//         "AssignedTo"
//       ).filter(`AssignedTo/Id eq ${currentUserId}`)();

//       const transformedTasks: Task[] = await Promise.all(
//         items.map(async (item: any, index: number) => {
//           try {
//             const creationitem = await spContext.web.lists.getByTitle("ProjectCreationList").items.getById(item.ProjectCreationListID.ID)
//               .select(
//                 "*",
//                 "ProjectType/ProjectType",
//                 "ProjectType/Id"
//               ).expand("ProjectType")();
//             console.log('Creation Item:', creationitem);

//             return {
//               sno: index + 1,
//               projectName: item.ProjectCreationListID?.ProjectName || "",
//               projectType: creationitem.ProjectType?.ProjectType || "",
//               deliverable: item.Deliverables || "",
//               deliverableId: item.Id || 0,
//               area: item.Area || "",
//               docType: item.DocumentType || "",
//               docNumber: item.DocNumber || "",
//               assignedTo: item.AssignedTo?.Title || "",
//               org: item.Organization || "",
//               status: item.Status,
//               projectId: item.ProjectCreationListID?.ID,
//               revisionnumber: item.RevisionNumber || "0",
//               creationDate: creationitem.Created || undefined,
//               clientName: creationitem.ClientName || "",
//             };
//           } catch (error) {
//             console.error(`Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`, error);
//             return {
//               sno: index + 1,
//               projectName: item.ProjectCreationListID?.ProjectName || "",
//               projectType: "Type A",
//               deliverable: item.Deliverables || "",
//               deliverableId: item.Id || 0,
//               area: item.Area || "",
//               docType: item.DocumentType || "",
//               docNumber: item.DocNumber || "",
//               assignedTo: item.AssignedTo?.Title || "",
//               org: item.Organization || "",
//               status: item.Status,
//               projectId: item.ProjectCreationListID?.ID,
//               revisionnumber: item.RevisionNumber || "0",
//               creationDate: undefined,
//               clientName: "",
//             };
//           }
//         })
//       );

//       setTasks(transformedTasks);
//       const pendingTasks = transformedTasks.filter(task => task.status === "Pending");
//       setFilteredTasks(pendingTasks);
//     } catch (error) {
//       console.error("Error loading tasks:", error);
//     }
//   };

//   const getDocumentComments = async (projectCreationID: number, deliverableDetailsID: number, revision: string) => {
//     try {
//       const items = await spContext.web.lists.getByTitle("DocumentComments").items
//         .select("*")
//         .filter(`ProjectID eq ${projectCreationID} and DeliverableDetailsID/ID eq ${deliverableDetailsID}`)
//         .orderBy("ID", false)();
//       if (items.length > 0) {
//         const comments: DocumentComment[] = items.map((item: any) => ({
//           id: item.Id,
//           userName: item.UserName || "",
//           commentDate: item.CommentDate ? new Date(item.CommentDate).toLocaleString('en-GB') : "",
//           pageNumber: item.PageNumber || "",
//           revision: item.Revision || "",
//           comment: item.Comment || ""
//         }));
//         setAllDocumentComments(comments);
//         console.log("All Document Comments:", comments);

//         const filteredComments = comments.filter(comment => comment.revision === revision);
//         setDocumentComments(filteredComments);

//         const uniqueVersions = [...new Set(comments.map(comment => comment.revision))].sort();
//         setVersionList(uniqueVersions);
//         setSelectedVersion(revision);
//         setShowDocumentComments(true);
//         console.log("Unique Versions:", uniqueVersions);
//         console.log("Filtered Document Comments:", filteredComments);
//       } else {
//         setShowDocumentComments(false);
//         setDocumentComments([]);
//         setAllDocumentComments([]);
//         setVersionList([]);
//       }
//     } catch (error) {
//       console.error("Error fetching document comments:", error);
//       setShowDocumentComments(false);
//     }
//   };

//   const onVersionChange = (version: string) => {
//     setSelectedVersion(version);
//     if (!version) {
//       setDocumentComments(allDocumentComments);
//     } else {
//       const filteredComments = allDocumentComments.filter(comment => comment.revision === version);
//       setDocumentComments(filteredComments);
//     }
//   };

//   const exportCommentsToExcel = () => {
//     const headers = ['Users', 'Comment Date', 'Page No.', 'Revision', 'Comments/Clarifications'];
//     const csvContent = [
//       headers.join(','),
//       ...documentComments.map(comment => [
//         `"${comment.userName}"`,
//         `"${comment.commentDate}"`,
//         `"${comment.pageNumber}"`,
//         `"${comment.revision}"`,
//         `"${comment.comment}"`
//       ].join(','))
//     ].join('\n');
//     const blob = new Blob([csvContent], { type: 'text/csv' });
//     const url = window.URL.createObjectURL(blob);
//     const a = document.createElement('a');
//     a.href = url;
//     a.download = `DocumentComments_${selectedTask?.docNumber || 'export'}.csv`;
//     document.body.appendChild(a);
//     a.click();
//     document.body.removeChild(a);
//     window.URL.revokeObjectURL(url);
//   };

//   const refreshDocComment = () => {
//     if (selectedTask) {
//       getDocumentComments(selectedTask.projectId, selectedTask.deliverableId, selectedTask.revisionnumber);
//     }
//   };

//   const pendingCount = tasks.filter(t => t.status === "Pending").length;
//   const completedCount = tasks.filter(t => t.status === "Approved" || t.status === "In-Progress").length;

//   const getStatusClass = (status: string) => {
//     switch (status) {
//       case "Pending": return styles.pendingStatus;
//       case "In-Progress": return styles.inProgressStatus;
//       case "Approved": return styles.completedStatus;
//       default: return styles.pendingStatus;
//     }
//   };

//   const getDeliverablesDetails = (status: string) => {
//     setCurrentFilter(status);
//     if (status === "All") {
//       setFilteredTasks(tasks);
//     } else if (status === "Completed" || status === "Approved") {
//       setFilteredTasks(tasks.filter(task => task.status === "Approved" || task.status === "In-Progress"));
//     } else {
//       setFilteredTasks(tasks.filter(task => task.status === status));
//     }
//   };

//   const getAuditHistoryDeliverables = async (item: any, context: WebPartContext): Promise<AuditHistoryItem[]> => {
//     const auditHistoryArr: AuditHistoryItem[] = [];

//     try {
//       const items = await spContext.web.lists.getByTitle("ProjectApprovals").items
//         .select(
//           "*",
//           "AssignedTo/Id",
//           "AssignedTo/Title",
//           "AssignedTo/EMail",
//           "DeliverablesDetailsId/Id",
//           "DeliverablesDetailsId/Deliverables",
//           "Author/Id",
//           "Author/Title",
//           "Author/EMail"
//         )
//         .expand(
//           "AssignedTo",
//           "DeliverablesDetailsId",
//           "Author"
//         )
//         .filter(`DeliverablesDetailsId/ID eq ${item.deliverableId}`)();

//       if (items.length > 0) {
//         items.forEach((auditHistoryItem: any, index: number) => {
//           let status = auditHistoryItem.Status;
//           if (status && status.toLowerCase() === 'pending') {
//             status = 'Pending';
//           }

//           auditHistoryArr.push({
//             sno: index + 1,
//             approvalLevel: auditHistoryItem.ApprovalLevel || `Level ${index}`,
//             assignedTo: auditHistoryItem.AssignedTo?.Title || 'N/A',
//             assignedToRole: auditHistoryItem.ApproverRole || '',
//             requestorName: auditHistoryItem.Author?.Title || auditHistoryItem.RequestorName || 'N/A',
//             requestedDate: auditHistoryItem.Created ? new Date(auditHistoryItem.Created).toLocaleString('en-GB') : 'N/A',
//             actionTakenBy: auditHistoryItem.ActionTakenBy?.Title || auditHistoryItem.ModifiedBy?.Title || 'N/A',
//             actionTakenOn: auditHistoryItem.Modified ? new Date(auditHistoryItem.Modified).toLocaleString('en-GB') : '',
//             remark: auditHistoryItem.Remarks || '',
//             status: status || 'Pending'
//           });
//         });
//       }
//     } catch (error) {
//       console.error("Error fetching audit history:", error);
//     }

//     return auditHistoryArr;
//   };

//   const handleViewClick = async (task: Task) => {
//     setSelectedTask(task);
//     setShowForm(true);
//     setUploadedFileName("");
//     setUploadedFileUrl("");
//     setComment("");
//     setExistingDocument(null);

//     try {
//       // Check if document already exists for this project and deliverable
//       const existingFiles = await spContext.web.lists.getByTitle("DeliverablesDocument").items
//         .select("ID", "FileLeafRef", "File/ServerRelativeUrl")
//         .expand("File")
//         .filter(`ProjectID eq '${task.projectId}' and DeliverablesDetailsId eq '${task.deliverableId}'`).orderBy("ID", false)();

//       if (existingFiles.length > 0) {
//         const file = existingFiles[0];
//         const serverRelativeUrl = file.File.ServerRelativeUrl;
//         const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;

//         setExistingDocument({
//           name: file.FileLeafRef,
//           url: fileUrl
//         });
//       }

//       // Load Audit History
//       const history = await getAuditHistoryDeliverables(task, props.context);
//       setAuditHistory(history);
//       setShowNoAuditHistory(history.length === 0);

//       // Load Document Comments
//       await getDocumentComments(task.projectId, task.deliverableId, task.revisionnumber);

//       // Load file and comment if applicable
//       if (task.status === "Approved" || task.status === "In-Progress") {
//         const deliverablesList = spContext.web.lists.getByTitle("DeliverablesDetails");
//         const deliverableItem = await deliverablesList.items
//           .getById(task.deliverableId)
//           .select("DeliverablesDocumentID/ID", "DocumentComments")
//           .expand("DeliverablesDocumentID")();

//         if (deliverableItem.DocumentComments) {
//           setComment(deliverableItem.DocumentComments);
//         }

//         const deliverablesDocumentId = deliverableItem.DeliverablesDocumentID?.ID;
//         if (deliverablesDocumentId) {
//           const fileItem = await spContext.web.lists.getByTitle("DeliverablesDocument").items
//             .getById(deliverablesDocumentId)
//             .select("ID", "FileLeafRef", "File/ServerRelativeUrl")
//             .expand("File")();

//           if (fileItem.File) {
//             setUploadedFileName(fileItem.FileLeafRef);
//             const serverRelativeUrl = fileItem.File.ServerRelativeUrl;
//             const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
//             setUploadedFileUrl(fileUrl);
//           }
//         }
//       } else {
//         setComment("");
//         setUploadedFileName("");
//         setUploadedFileUrl("");
//       }
//     } catch (error) {
//       console.error("Error loading task details:", error);
//       setAuditHistory([]);
//       setShowNoAuditHistory(true);
//       setShowDocumentComments(false);
//     }
//   };

//   const handleBackClick = () => {
//     setSelectedTask(null);
//     setShowForm(false);
//     setAuditHistory([]);
//     setShowNoAuditHistory(false);
//     setShowDocumentComments(false);
//     setDocumentComments([]);
//     setExistingDocument(null);
//   };

//   const handleSubmitClick = async () => {
//     if (!selectedTask || !selectedFile) {
//       setPopup({
//         isOpen: true,
//         type: 'validation',
//         title: 'Validation',
//         message: 'Please select a file before submitting.'
//       });
//       return;
//     }

//     if (!documentControllerId) {
//       setPopup({
//         isOpen: true,
//         type: 'error',
//         title: 'Configuration Missing',
//         message: 'Document Controller not configured. Please contact administrator.'
//       });
//       return;
//     }

//     try {
//       // Calculate new revision (current + 1)
//       const currentRevision = parseInt(selectedTask.revisionnumber || "0");

//       //  this was used before and it was working
//       // const newRevision = (currentRevision + 1).toString();

//       const newRevision =
//         existingDocument && selectedTask?.status === "Pending"
//           ? (currentRevision + 1).toString()
//           : currentRevision.toString();
//       // STEP 1: Upload file to DeliverablesDocument library
//       const folder = spContext.web.getFolderByServerRelativePath("DeliverablesDocument");
//       const uploadResult = await folder.files.addUsingPath(selectedFile.name, selectedFile, { Overwrite: true });

//       // Get list item associated with uploaded file
//       const fileItem = await uploadResult.file.getItem();
//       const uploadedFileItemId = (fileItem as any).Id;
//       console.log("Uploaded File Item ID:", uploadedFileItemId);

//       // Update the uploaded file's metadata in the document library
//       let deliverableIdStr = selectedTask.deliverableId.toString();
//       let projectIdStr = selectedTask.projectId.toString();

//       await spContext.web.lists.getByTitle("DeliverablesDocument").items
//         .getById(uploadedFileItemId)
//         .update({
//           ProjectID: projectIdStr,
//           DeliverablesDetailsId: deliverableIdStr,
//           Revision: newRevision
//         });

//       console.log("File metadata updated successfully");

//       // STEP 2: Update DeliverablesDetails item with new revision
//       await spContext.web.lists.getByTitle("DeliverablesDetails").items
//         .getById(selectedTask.deliverableId)
//         .update({
//           DocumentComments: comment,
//           RevisionNumber: newRevision,
//           Status: "In-Progress",
//           DeliverablesDocumentIDId: uploadedFileItemId
//         });

//       console.log("DeliverablesDetails updated successfully");

//       // STEP 3: Mark Vendor's approval as Completed
//       const projectId = selectedTask?.projectId;
//       const deliverableId = selectedTask?.deliverableId;
//       const approverType = "Vendor";
//       const currentUser = await sp.web.currentUser();

//       try {
//         const items = await spContext.web.lists
//           .getByTitle("ProjectApprovals")
//           .items.select(
//             "*",
//             "DeliverablesDetailsId/ID",
//             "ProjectCreationListID/ID",
//             "AssignedTo/ID",
//             "AssignedTo/Title",
//             "AssignedTo/EMail"
//           )
//           .expand("DeliverablesDetailsId", "ProjectCreationListID", "AssignedTo")
//           .orderBy("SerialNumber", true)
//           .filter(
//             `ProjectCreationListID/ID eq ${projectId} and ` +
//             `DeliverablesDetailsId/ID eq ${deliverableId} and ` +
//             `AssignedTo/ID eq ${currentUser.Id} and ` +
//             `ApproverRole eq '${approverType}' and ` +
//             `Status eq 'Pending'`
//           )();

//         if (items.length > 0) {
//           const vendorTaskId = items[0].Id;
//           await spContext.web.lists.getByTitle("ProjectApprovals").items.getById(vendorTaskId).update({
//             Status: "Completed",
//             Remarks: comment,
//             ApprovalDate: new Date()
//           });
//           console.log("Vendor task marked as completed");
//         } else {
//           console.log("No matching ProjectApprovals found for current user");
//         }
//       } catch (error) {
//         console.error("Error updating ProjectApprovals:", error);
//       }

//       // STEP 4: Create new ProjectApproval item for Document Controller with new revision
//       await spContext.web.lists.getByTitle("ProjectApprovals").items.add({
//         DeliverablesDetailsIdId: selectedTask.deliverableId,
//         ProjectCreationListIDId: selectedTask.projectId,
//         AssignedToId: documentControllerId,
//         DocumentType: selectedTask.docType,
//         ApproverRole: "Document Controller",
//         ProjectType: selectedTask.projectType,
//         Level: "Level 1",
//         SerialNumber: 0,
//         ApprovalCriteria: "Anyone",
//         RequestedById: currentUser.Id,
//         RequestedDate: new Date(),
//         IncomingDate: new Date(),
//         RequestedRole: "Vendor",
//         RevisionNumber: newRevision,
//         Status: "Pending"
//       });

//       setPopup({
//         isOpen: true,
//         type: 'success',
//         title: 'Success',
//         message: 'Task submitted successfully.',
//         onConfirm: () => {
//           setPopup(prev => ({ ...prev, isOpen: false }));
//           handleBackClick();
//           window.location.reload();
//         }
//       });

//     } catch (error) {
//       console.error("Error in submission:", error);
//       setPopup({
//         isOpen: true,
//         type: 'error',
//         title: 'Error',
//         message: 'An error occurred during submission. Please check console for details.'
//       });
//     }
//   }

//   return (
//       <div id="wrapper" ref={elementRef}>
//           <div className="app-menu" id="myHeader">
//             <VerticalSideBar _context={sp} />
//           </div>

//           <div className="content-page">
//             <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
//       <div
//           className="content"
//           style={{
//             marginLeft: `${!useHide ? "240px" : "80px"}`,

//             marginTop: "0rem",
//           }}
//         >
//           <div className="container-fluid  paddb">
//     <div className={styles.myTask}>
//       {!showForm ? (
//         <>
//           <h2 className='fw-bold text-dark header-title mt-2'>My Tasks</h2>

//           {/* Tiles */}
//           <div className={styles.tilesContainer}>
//             <div
//               className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ''}`}
//               onClick={() => getDeliverablesDetails("Pending")}
//             >
//               <div className={styles.tileBody}>
//                 <div className={styles.tileRow}>
//                   <div className={styles.tileIcon}>
//                     <div className={`${styles.avatar} ${styles.primary}`}>
//                       <span className={styles.icon}>⏱</span>
//                     </div>
//                   </div>
//                   <div className={styles.tileContent}>
//                     <h3 className={styles.tileCount}>{pendingCount}</h3>
//                     <p className={styles.tileLabel}>Pending</p>
//                   </div>
//                 </div>
//               </div>
//             </div>

//             <div
//               className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ''}`}
//               onClick={() => getDeliverablesDetails("Approved")}
//             >
//               <div className={styles.tileBody}>
//                 <div className={styles.tileRow}>
//                   <div className={styles.tileIcon}>
//                     <div className={`${styles.avatar} ${styles.success}`}>
//                       <span className={styles.icon}>✔️</span>
//                     </div>
//                   </div>
//                   <div className={styles.tileContent}>
//                     <h3 className={styles.tileCount}>{completedCount}</h3>
//                     <p className={styles.tileLabel}>Completed</p>
//                   </div>
//                 </div>
//               </div>
//             </div>
//           </div>

//           {/* Table */}
//           <div style={{clear:'both', marginTop:'15px'}} className='card mt-2'>

//             <div className='card-body'>
//             <div className={styles.mainTableContainer}>
//             <div className={styles.tableCard}>
//               <div className={styles.tableWrapper}>
//                 <table className={styles.taskTable}>
//                   <thead>
//                     <tr>
//                       <th style={{minWidth:'70px'}}>S.No</th>
//                       <th>Project Name</th>
//                       <th>Project Type</th>
//                       <th>Deliverable</th>
//                       <th>Document Type</th>
//                       <th>Assigned To</th>
//                       <th>Organization</th>
//                       <th style={{minWidth:'90px'}}>Status</th>
//                       <th style={{minWidth:'70px'}}>Action</th>
//                     </tr>
//                   </thead>
//                   <tbody>
//                     {filteredTasks.map((task, index) => (
//                       <tr key={task.sno}>
//                         <td style={{minWidth:'70px'}}>{index + 1}</td>
//                         <td>{task.projectName}</td>
//                         <td>{task.projectType}</td>
//                         <td>{task.deliverable}</td>
//                         <td>{task.docType}</td>
//                         <td>{task.assignedTo}</td>
//                         <td>{task.org}</td>
//                         <td style={{minWidth:'90px'}}>
//                           <span className={`${styles.statusBadge} ${getStatusClass(task.status)}`}>
//                             {task.status}
//                           </span>
//                         </td>
//                         <td style={{minWidth:'70px'}}>
//                           <span
//                             className={styles.actionIcon}
//                             onClick={() => handleViewClick(task)}
//                             title={`View ${task.deliverable}`}
//                           >
//                             👁️
//                           </span>
//                         </td>
//                       </tr>
//                     ))}
//                   </tbody>
//                 </table>
//               </div>
//             </div>
//             {/* <div className={styles.scrollHint}><em>Scroll horizontally → to view all columns</em></div> */}
//           </div>
//             </div>
//           </div>

//         </>
//       ) : (
//         <>
//           {/* Form View */}
//           <div className={styles.formContainer}>
//             <div className={styles.formHeader}>
//               <h3 style={{margin:'0px'}}>
//                 My Task &gt;&gt; {selectedTask?.docNumber}
//               </h3>
//               <button type="button" className={styles.backButton} onClick={handleBackClick}>Back</button>
//             </div>
//             <div className={styles.formView}>
//               <div className={styles.formInner}>
//                 <div className={styles.formBody}>
//                   <div className={styles.formGrid}>
//                     <div>
//                       <label>Project Name</label>
//                       <input type="text" value={selectedTask?.projectName || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Project Type</label>
//                       <input type="text" value={selectedTask?.projectType || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Client Name</label>
//                       <input type="text" value={selectedTask?.clientName || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Project Date</label>
//                       <input type="text" value={selectedTask?.creationDate ? new Date(selectedTask.creationDate).toLocaleDateString('en-GB') : ''} disabled />
//                     </div>
//                     <div>
//                       <label>Prepared By</label>
//                       <input type="text" value={selectedTask?.assignedTo || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Deliverable</label>
//                       <input type="text" value={selectedTask?.deliverable || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Document Type</label>
//                       <input type="text" value={selectedTask?.docType || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Document Number</label>
//                       <input
//                         type="text"
//                         value={selectedTask?.docNumber || ''}
//                         disabled
//                       />
//                     </div>
//                     <div>
//                       <label>Area</label>
//                       <input type="text" value={selectedTask?.area || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Organization</label>
//                       <input type="text" value={selectedTask?.org || ''} disabled />
//                     </div>
//                     <div>
//                       <label>Revision Number</label>
//                       <input
//                         type="text"
//                         value={
//                           selectedTask?.status === "Pending" && existingDocument
//                             ? (parseInt(selectedTask?.revisionnumber || "0") + 1).toString()
//                             : selectedTask?.revisionnumber || "0"
//                         }
//                         disabled
//                       />
//                     </div>
//                   </div>

//                   {/* Upload & Comment Section */}
//                   {/* Upload & Comment Section */}
//                   <div className={styles.formActions}>
//                     {/* For Pending tasks with existing document */}
//                     {selectedTask?.status === "Pending" && existingDocument && (
//                       <>
//                         <label>Uploaded Document:</label>
//                         <p>
//                           <a href={existingDocument.url} target="_blank" rel="noopener noreferrer">
//                             {existingDocument.name}
//                           </a>
//                         </p>
//                         <label>Upload Document*</label>
//                         <input type="file" onChange={(e) => setSelectedFile(e.target.files?.[0] || null)} />
//                       </>
//                     )}

//                     {/* For Completed/In-Progress tasks - show only uploaded file */}
//                     {(selectedTask?.status === "Approved" || selectedTask?.status === "In-Progress") && uploadedFileName && (
//                       <>
//                         <label>Uploaded Document:</label>
//                         <p>
//                           <a href={uploadedFileUrl} target="_blank" rel="noopener noreferrer">
//                             {uploadedFileName}
//                           </a>
//                         </p>
//                       </>
//                     )}

//                     {/* Show upload field only for Pending tasks without existing document */}
//                     {selectedTask?.status === "Pending" && !existingDocument && (
//                       <>
//                         <label>Upload Document*</label>
//                         <input type="file" onChange={(e) => setSelectedFile(e.target.files?.[0] || null)} />
//                       </>
//                     )}

//                     <label>Comment*</label>
//                     <textarea
//                       placeholder="Enter your comment"
//                       value={comment}
//                       onChange={(e) => setComment(e.target.value)}
//                       disabled={selectedTask?.status === "Approved" || selectedTask?.status === "In-Progress"}
//                     />
//                   </div>

//                   {/* Document Comments Accordion */}
//                   {showDocumentComments && (
//                     <div className={styles.accordionItem}>
//                       <h2 className={styles.accordionHeader}>
//                         <div className={styles.accordionButton}>
//                           Document Comments
//                         </div>
//                       </h2>

//                       <div className={styles.accordionBody}>
//                         <div className={styles.customCard}>
//                           <div className={styles.documentCommentsHeader}>
//                             <select
//                               className={styles.formSelect}
//                               value={selectedVersion}
//                               onChange={(e) => onVersionChange(e.target.value)}
//                             >
//                               <option value="">-- Select Revision --</option>
//                               {versionList.map(version => (
//                                 <option key={version} value={version}>{version}</option>
//                               ))}
//                             </select>
//                             <button
//                               className={styles.btnOutlineSuccess}
//                               type="button"
//                               onClick={exportCommentsToExcel}
//                             >
//                               Export to Excel
//                             </button>
//                             <button
//                               className={styles.btnOutlineSuccess}
//                               type="button"
//                               onClick={refreshDocComment}
//                             >
//                               ↻
//                             </button>
//                           </div>

//                           <div className={styles.ribbonContent}>
//                             <table className={styles.commentsTable}>
//                               <thead>
//                                 <tr>
//                                   <th style={{ minWidth: '80px', maxWidth: '80px' }}>Users</th>
//                                   <th style={{ minWidth: '100px', maxWidth: '100px' }}>Comment Date</th>
//                                   <th style={{ minWidth: '80px', maxWidth: '80px' }}>Page No.</th>
//                                   <th style={{ minWidth: '80px', maxWidth: '80px' }}>Revision</th>
//                                   <th style={{ minWidth: '200px', maxWidth: '200px' }}>Comments</th>
//                                 </tr>
//                               </thead>
//                               <tbody style={{ maxHeight: '250px', overflowY: 'auto' }}>
//                                 {documentComments.map((commentItem) => (
//                                   <tr key={commentItem.id}>
//                                     <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
//                                       {commentItem.userName}
//                                     </td>
//                                     <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '100px', maxWidth: '100px' }}>
//                                       {commentItem.commentDate}
//                                     </td>
//                                     <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
//                                       {commentItem.pageNumber}
//                                     </td>
//                                     <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
//                                       {commentItem.revision}
//                                     </td>
//                                     <td style={{ padding: '15px', verticalAlign: 'top', minWidth: '200px', maxWidth: '200px' }}>
//                                       {commentItem.comment}
//                                     </td>
//                                   </tr>
//                                 ))}
//                               </tbody>
//                             </table>
//                           </div>
//                         </div>
//                       </div>
//                     </div>
//                   )}

//                   {/* Audit History */}
//                   <div className={styles.auditHistory}>
//                     <h4>Audit History</h4>
//                     {showNoAuditHistory && auditHistory.length === 0 ? (
//                       <p>No audit history available</p>
//                     ) : (
//                       <table className='mtablemyt'>
//                         <thead>
//                           <tr>
//                             <th style={{minWidth:'70px',maxWidth:'70px'}}>SNo</th>
//                             <th>Approval Level</th>
//                             <th>Assigned To</th>
//                             <th>Assigned To Role</th>
//                             <th>Requestor Name</th>
//                             <th>Requested Date</th>
//                             <th>Action Taken By</th>
//                             <th>Action Taken On</th>
//                             <th>Remark</th>
//                             <th style={{minWidth:'110px',maxWidth:'110px'}}>Status</th>
//                           </tr>
//                         </thead>
//                         <tbody>
//                           {auditHistory.map((historyItem) => (
//                             <tr key={historyItem.sno}>
//                               <td style={{minWidth:'70px',maxWidth:'70px'}}>{historyItem.sno}</td>
//                               <td>{historyItem.approvalLevel}</td>
//                               <td>{historyItem.assignedTo}</td>
//                               <td>{historyItem.assignedToRole}</td>
//                               <td>{historyItem.requestorName}</td>
//                               <td>{historyItem.requestedDate}</td>
//                               <td>{historyItem.assignedTo}</td>
//                               <td>{historyItem.actionTakenOn}</td>
//                               <td>{historyItem.remark}</td>
//                               <td style={{minWidth:'110px',maxWidth:'110px'}}>
//                                 <span className={`${styles.statusBadge} ${getStatusClass(historyItem.status)}`}>
//                                   {historyItem.status}
//                                 </span>
//                               </td>
//                             </tr>
//                           ))}
//                         </tbody>
//                       </table>
//                     )}
//                   </div>

//                   <div className={styles.formButtons}>
//                     {selectedTask?.status === "Pending" && (
//                       <>
//                         <button type="button" className={styles.submitButton} onClick={handleSubmitClick}>Submit</button>
//                         <button type="button" className={styles.cancelButton} onClick={handleBackClick}>Cancel</button>
//                       </>
//                     )}
//                     {selectedTask?.status === "Approved" && (
//                       <button type="button" className={styles.cancelButton} onClick={handleBackClick}>Back</button>
//                     )}
//                   </div>
//                 </div>
//               </div>
//             </div>
//           </div>
//         </>
//       )}
//       <CustomPopup
//         isOpen={popup.isOpen}
//         type={popup.type}
//         title={popup.title}
//         message={popup.message}
//         onConfirm={popup.onConfirm}
//         onCancel={() => setPopup(prev => ({ ...prev, isOpen: false }))}
//         onClose={() => setPopup(prev => ({ ...prev, isOpen: false }))}
//         onSuccessOk={() => {
//           setPopup(prev => ({ ...prev, isOpen: false }));
//           if (popup.onConfirm) popup.onConfirm();
//         }}
//       />
//     </div>
//     </div></div>
//           </div>
//       </div>
//   );
// };

// export default MyTask;





import * as React from 'react';
import styles from './MyTasks.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/web";
import CustomPopup from '../../myProject/components/CustomPopup';
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
import UserContext from "../../../GlobalContext/context";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../loc/pnpjsConfig";
import loaderGif from "../assets/Loder.gif";
import totalim from "../assets/total.png";
import approve from "../assets/approve.png";
import pen from "../assets/pen.png";
import newlogo from "../assets/logo-high.png";
import eye from "../assets/eye.png";


interface IMyTaskProps {
  context: WebPartContext;
  siteUrl: string;
  // description: string;
  // siteUrl: string;
  // userDisplayName: string;
  // isDarkTheme: boolean;
  // environmentMessage: string;
  // hasTeamsContext: boolean;
}

type ActionType =
  | "Submitted"

const actionConfig: Record<ActionType, { subject: string; message: string }> = {
  Submitted: {
    subject: "Approval Action Assigned",
    message: "The assigned task has been completed and the document has been submitted for your review. Kindly review the submitted document and proceed with the next steps."
  }
};

interface Task {
  projectId: any;
  sno: number;
  projectName: string;
  projectType: string;
  deliverable: string;
  deliverableId: number;
  area: string;
  docType: string;
  docNumber: string;
  revisionnumber: string;
  assignedTo: string;
  preparedBy: string;
  preparedByEmail?: string;
  IsReworked: string;
  org: string;
  status: "Pending" | "Approved" | "In-Progress";
  clientName: string;
  creationDate: Date | undefined;
  DueDate?: Date | undefined;
}

interface AuditHistoryItem {
  sno: number;
  approvalLevel: string;
  assignedTo: string;
  assignedToRole: string;
  requestorName: string;
  requestedDate: string;
  actionTakenBy: string;
  actionTakenOn: string;
  remark: string;
  status: string;
}

interface DocumentComment {
  id: number;
  userName: string;
  commentDate: string;
  pageNumber: string;
  revision: string;
  comment: string;
}
interface UploadedFile {
  name: string;
  file: File | null;
  url?: string;
  id?: number;
}
const MyTask = ({ props }: any) => {
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP();
  const siteUrl = props.siteUrl
  const { useHide }: any = React.useContext(UserContext);
  const [tasks, setTasks] = React.useState<Task[]>([]);
  const [currentFilter, setCurrentFilter] = React.useState<string>("Pending");
  const [filteredTasks, setFilteredTasks] = React.useState<Task[]>([]);
  const [selectedTask, setSelectedTask] = React.useState<Task | null>(null);
  const [showForm, setShowForm] = React.useState<boolean>(false);
  const [auditHistory, setAuditHistory] = React.useState<AuditHistoryItem[]>([]);
  const [showNoAuditHistory, setShowNoAuditHistory] = React.useState<boolean>(false);
  //const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [selectedFiles, setSelectedFiles] = React.useState<UploadedFile[]>([]);
  const [showSubmitLoader, setShowSubmitLoader] = React.useState<boolean>(false);


  const [comment, setComment] = React.useState<string>("");
  const [documentControllerId, setDocumentControllerId] = React.useState<number | null>(null);
  const [documentControllerName, setDocumentControllerName] = React.useState<string | null>(null);
  const [uploadedFileName, setUploadedFileName] = React.useState<string>("");
  const [uploadedFileUrl, setUploadedFileUrl] = React.useState<string>("");
  const [documentComments, setDocumentComments] = React.useState<DocumentComment[]>([]);
  const [allDocumentComments, setAllDocumentComments] = React.useState<DocumentComment[]>([]);
  const [versionList, setVersionList] = React.useState<string[]>([]);
  const [selectedVersion, setSelectedVersion] = React.useState<string>("");
  const [showDocumentComments, setShowDocumentComments] = React.useState<boolean>(false);
  // const [existingDocument, setExistingDocument] = React.useState<{
  //   name: string;
  //   url: string;
  // } | null>(null);
  const [existingDocuments, setExistingDocuments] = React.useState<Array<{
    name: string;
    url: string;
    id: number;
    revision?: string;  // Optional बना दें
    created?: string;   // Optional बना दें
  }>>([]);

  const [popup, setPopup] = React.useState<{
    isOpen: boolean;
    type: 'confirmation' | 'validation' | 'success' | 'error';
    title: string;
    message: string;
    onConfirm?: () => void;
  }>({
    isOpen: false,
    type: 'success',
    title: '',
    message: ''
  });

  const spContext = spfi().using(SPFx(props.context));

  const getDocumentController = async (): Promise<number | null> => {
    try {
      const items = await spContext.web.lists.getByTitle("ProjectConfiguration").items
        .select(
          "*",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DocumentController/ID"
        )
        .expand("DocumentController")
        .orderBy("Created", false)()

      if (items.length > 0) {
        const documentControllerId = items[0].DocumentControllerId;
        const documentControllerName = items[0].DocumentController?.Title;
        console.log("Document Controller ID:", documentControllerId);
        return documentControllerId;
      }
      return null;
    } catch (error) {
      console.error("Error fetching Document Controller:", error);
      return null;
    }
  };
  const getDocumentControllerName = async (): Promise<string | null> => {
    try {
      const items = await spContext.web.lists.getByTitle("ProjectConfiguration").items
        .select(
          "*",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DocumentController/ID"
        )
        .expand("DocumentController")
        .orderBy("Created", false)()

      if (items.length > 0) {
        const documentControllerName = items[0].DocumentController?.Title;
        console.log("Document Controller ID:", documentControllerName);
        return documentControllerName;
      }
      return null;
    } catch (error) {
      console.error("Error fetching Document Controller:", error);
      return null;
    }
  };

  React.useEffect(() => {
    const initializeData = async () => {
      const docControllerId = await getDocumentController();
      const documentControllerName = await getDocumentControllerName();
      setDocumentControllerId(docControllerId);
      setDocumentControllerName(documentControllerName);
      await fetchTasks();
      setTimeout(() => {
        const menu = document.querySelector('.sidebar-openBa');
        if (menu instanceof HTMLElement) menu.style.display = 'none';

        const search = document.querySelector('.search_bar');
        if (search instanceof HTMLElement) search.style.display = 'none';
      }, 1000);
    };

    initializeData();
  }, []);

  const fetchTasks = async () => {
    const currentUserId = (await spContext.web.currentUser()).Id;
    try {
      const items = await spContext.web.lists.getByTitle("DeliverablesDetails").items.select(
        "*",
        "ProjectCreationListID/ID",
        "ProjectCreationListID/ProjectName",
        "AssignedTo/Id",
        "AssignedTo/Title"
      ).expand(
        "ProjectCreationListID",
        "AssignedTo"
      ).filter(`AssignedTo/Id eq ${currentUserId}`)();

      const transformedTasks: Task[] = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            const creationitem = await spContext.web.lists.getByTitle("ProjectCreationList").items.getById(item.ProjectCreationListID.ID)
              .select(
                "*",
                "ProjectType/ProjectType",
                "ProjectType/Id",
                "PreparedBy/Id",
                "PreparedBy/Title",
                "PreparedBy/EMail",
              ).expand("ProjectType", "PreparedBy")();
            console.log('Creation Item:', creationitem);
            console.log('Creation Item PreparedBy:', creationitem.PreparedBy);

            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: creationitem.ProjectType?.ProjectType || "",
              deliverable: item.Deliverables || "",
              deliverableId: item.Id || 0,
              area: item.Area || "",
              docType: item.DocumentType || "",
              docNumber: item.DocNumber || "",
              DueDate: item.DueDate || undefined,
              assignedTo: item.AssignedTo?.Title || "",
              IsReworked: item.IsReworked || "",
              preparedBy: creationitem.PreparedBy?.Title || "",
              preparedByEmail: creationitem.PreparedBy?.EMail || "",
              org: item.Organization || "",
              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: creationitem.Created || undefined,
              clientName: creationitem.ClientName || "",
            };
          } catch (error) {
            console.error(`Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`, error);
            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: "Type A",
              deliverable: item.Deliverables || "",
              deliverableId: item.Id || 0,
              area: item.Area || "",
              docType: item.DocumentType || "",
              docNumber: item.DocNumber || "",
              assignedTo: item.AssignedTo?.Title || "",
              IsReworked: item.IsReworked || "",
              preparedBy: "",
              preparedByEmail: "",
              org: item.Organization || "",
              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: undefined,
              clientName: "",
            };
          }
        })
      );

      setTasks(transformedTasks);
      const pendingTasks = transformedTasks.filter(task => task.status === "Pending");
      setFilteredTasks(pendingTasks);
    } catch (error) {
      console.error("Error loading tasks:", error);
    }
  };

  const getDocumentComments = async (projectCreationID: number, deliverableDetailsID: number, revision: string) => {
    try {
      const items = await spContext.web.lists.getByTitle("DocumentComments").items
        .select("*")
        .filter(`ProjectID eq ${projectCreationID} and DeliverableDetailsID/ID eq ${deliverableDetailsID}`)
        .orderBy("ID", false)();
      if (items.length > 0) {
        const comments: DocumentComment[] = items.map((item: any) => ({
          id: item.Id,
          userName: item.UserName || "",
          commentDate: item.CommentDate ? new Date(item.CommentDate).toLocaleString('en-GB') : "",
          pageNumber: item.PageNumber || "",
          revision: item.Revision || "",
          comment: item.Comment || ""
        }));
        setAllDocumentComments(comments);
        console.log("All Document Comments:", comments);

        const filteredComments = comments.filter(comment => comment.revision === revision);
        setDocumentComments(filteredComments);

        const uniqueVersions = [...new Set(comments.map(comment => comment.revision))].sort();
        setVersionList(uniqueVersions);
        setSelectedVersion(revision);
        setShowDocumentComments(true);
        console.log("Unique Versions:", uniqueVersions);
        console.log("Filtered Document Comments:", filteredComments);
      } else {
        setShowDocumentComments(false);
        setDocumentComments([]);
        setAllDocumentComments([]);
        setVersionList([]);
      }
    } catch (error) {
      console.error("Error fetching document comments:", error);
      setShowDocumentComments(false);
    }
  };

  const onVersionChange = (version: string) => {
    setSelectedVersion(version);
    if (!version) {
      setDocumentComments(allDocumentComments);
    } else {
      const filteredComments = allDocumentComments.filter(comment => comment.revision === version);
      setDocumentComments(filteredComments);
    }
  };

  const exportCommentsToExcel = () => {
    const headers = ['Users', 'Comment Date', 'Page No.', 'Revision', 'Comments/Clarifications'];
    const csvContent = [
      headers.join(','),
      ...documentComments.map(comment => [
        `"${comment.userName}"`,
        `"${comment.commentDate}"`,
        `"${comment.pageNumber}"`,
        `"${comment.revision}"`,
        `"${comment.comment}"`
      ].join(','))
    ].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `DocumentComments_${selectedTask?.docNumber || 'export'}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  const refreshDocComment = () => {
    if (selectedTask) {
      getDocumentComments(selectedTask.projectId, selectedTask.deliverableId, selectedTask.revisionnumber);
    }
  };

  const pendingCount = tasks.filter(t => t.status === "Pending").length;
  const completedCount = tasks.filter(t => t.status === "Approved" || t.status === "In-Progress").length;
  const TotalDocumentCount = tasks.length;

  const getStatusClass = (status: string) => {
    switch (status) {
      case "Pending": return styles.pendingStatus;
      case "In-Progress": return styles.inProgressStatus;
      case "Approved": return styles.completedStatus;
      default: return styles.pendingStatus;
    }
  };




  const getDeliverablesDetails = (status: string) => {
    setCurrentFilter(status);
    if (status === "All") {
      setFilteredTasks(tasks);
    } else if (status === "Completed" || status === "Approved") {
      setFilteredTasks(tasks.filter(task => task.status === "Approved" || task.status === "In-Progress"));
    } else {
      setFilteredTasks(tasks.filter(task => task.status === status));
    }
  };

  const getAuditHistoryDeliverables = async (item: any, context: WebPartContext): Promise<AuditHistoryItem[]> => {
    const auditHistoryArr: AuditHistoryItem[] = [];

    try {
      const items = await spContext.web.lists.getByTitle("ProjectApprovals").items
        .select(
          "*",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "DeliverablesDetailsId/Id",
          "DeliverablesDetailsId/Deliverables",
          "Author/Id",
          "Author/Title",
          "Author/EMail"
        )
        .expand(
          "AssignedTo",
          "DeliverablesDetailsId",
          "Author"
        )
        .filter(`DeliverablesDetailsId/ID eq ${item.deliverableId}`)();

      if (items.length > 0) {
        items.forEach((auditHistoryItem: any, index: number) => {
          let status = auditHistoryItem.Status;
          if (status && status.toLowerCase() === 'pending') {
            status = 'Pending';
          }

          auditHistoryArr.push({
            sno: index + 1,
            approvalLevel: auditHistoryItem.ApprovalLevel || `Level ${index}`,
            assignedTo: auditHistoryItem.AssignedTo?.Title || 'N/A',
            assignedToRole: auditHistoryItem.ApproverRole || '',
            requestorName: auditHistoryItem.Author?.Title || auditHistoryItem.RequestorName || 'N/A',
            requestedDate: auditHistoryItem.Created ? new Date(auditHistoryItem.Created).toLocaleString('en-GB') : 'N/A',
            actionTakenBy: auditHistoryItem.ActionTakenBy?.Title || auditHistoryItem.ModifiedBy?.Title || 'N/A',
            actionTakenOn: auditHistoryItem.Modified ? new Date(auditHistoryItem.Modified).toLocaleString('en-GB') : '',
            remark: auditHistoryItem.Remarks || '',
            status: status || 'Pending'
          });
        });
      }
    } catch (error) {
      console.error("Error fetching audit history:", error);
    }

    return auditHistoryArr;
  };

  const handleViewClick = async (task: Task) => {
    setSelectedTask(task);
    setShowForm(true);
    setUploadedFileName("");
    setUploadedFileUrl("");
    setComment("");
    setSelectedFiles([]);
    setExistingDocuments([]);

    try {
      // Check if document already exists for this project and deliverable
      // Check ALL documents for this project and deliverable
      // const allFiles = await spContext.web.lists.getByTitle("DeliverablesDocument").items
      //   .select("ID", "FileLeafRef", "File/ServerRelativeUrl", "Revision", "Created")
      //   .expand("File")
      //   .filter(`ProjectID eq '${task.projectId}' and DeliverablesDetailsId eq '${task.deliverableId}'`)
      //   .orderBy("Created", false)();
      const allFiles = await spContext.web.lists
        .getByTitle("DeliverablesDocument")
        .items
        .select("ID", "FileLeafRef", "File/ServerRelativeUrl", "Revision", "Created")
        .expand("File")
        .filter(
          `ProjectID eq '${task.projectId}' and DeliverablesDetailsId eq '${task.deliverableId}' and Revision eq '${task.revisionnumber.toString()}'`
        )
        .orderBy("Created", false)();


      if (allFiles.length > 0) {
        const documents = allFiles.map(file => {
          const serverRelativeUrl = file.File.ServerRelativeUrl;
          const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;

          return {
            name: file.FileLeafRef,
            url: fileUrl,
            id: file.ID,
            revision: file.Revision,
            created: file.Created
          };
        });

        setExistingDocuments(documents);
      }

      // Load Audit History
      const history = await getAuditHistoryDeliverables(task, props.context);
      setAuditHistory(history);
      setShowNoAuditHistory(history.length === 0);

      // Load Document Comments
      await getDocumentComments(task.projectId, task.deliverableId, task.revisionnumber);

      // Load file and comment if applicable
      if (task.status === "Approved" || task.status === "In-Progress") {
        const deliverablesList = spContext.web.lists.getByTitle("DeliverablesDetails");
        const deliverableItem = await deliverablesList.items
          .getById(task.deliverableId)
          .select("DeliverablesDocumentID/ID", "DocumentComments")
          .expand("DeliverablesDocumentID")();

        if (deliverableItem.DocumentComments) {
          setComment(deliverableItem.DocumentComments);
        }

        const deliverablesDocumentId = deliverableItem.DeliverablesDocumentID?.ID;
        if (deliverablesDocumentId) {
          const fileItem = await spContext.web.lists.getByTitle("DeliverablesDocument").items
            .getById(deliverablesDocumentId)
            .select("ID", "FileLeafRef", "File/ServerRelativeUrl")
            .expand("File")();

          if (fileItem.File) {
            setUploadedFileName(fileItem.FileLeafRef);
            const serverRelativeUrl = fileItem.File.ServerRelativeUrl;
            const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
            setUploadedFileUrl(fileUrl);
          }
        }
      } else {
        setComment("");
        setUploadedFileName("");
        setUploadedFileUrl("");
      }
    } catch (error) {
      console.error("Error loading task details:", error);
      setAuditHistory([]);
      setShowNoAuditHistory(true);
      setShowDocumentComments(false);
    }
  };

  const handleBackClick = () => {
    setSelectedTask(null);
    setShowForm(false);
    setAuditHistory([]);
    setShowNoAuditHistory(false);
    setShowDocumentComments(false);
    setDocumentComments([]);
    setExistingDocuments([]); // Empty array with proper type
  };

  const handleSubmitClick = async () => {
    setShowSubmitLoader(true); // ✅ ADD HERE
    if (!selectedTask || selectedFiles.length === 0) {
      setPopup({
        isOpen: true,
        type: 'validation',
        title: 'Validation',
        message: 'Please select at least one file before submitting.'
      });
      return;
    }

    if (!documentControllerId) {
      setPopup({
        isOpen: true,
        type: 'error',
        title: 'Configuration Missing',
        message: 'Document Controller not configured. Please contact administrator.'
      });
      return;
    }

    try {
      // Calculate new revision (current + 1)
      // const currentRevision = parseInt(selectedTask.revisionnumber || "0");

      //  this was used before and it was working
      // const newRevision = (currentRevision + 1).toString();

      // const newRevision =
      //   existingDocuments && selectedTask?.status === "Pending"
      //     ? (currentRevision + 1).toString()
      //     : currentRevision.toString();
      const currentRevision = parseInt(selectedTask.revisionnumber || "0");

      // Only increment if IsReworked is "Yes"
      const newRevision = selectedTask.IsReworked === "Yes"
        ? (currentRevision + 1).toString()
        : currentRevision.toString();


      // STEP 1: Upload file to DeliverablesDocument library
      const uploadedFileIds: number[] = [];

      // Upload all selected files
      for (const fileObj of selectedFiles) {
        if (fileObj.file) {
          const folder = spContext.web.getFolderByServerRelativePath("DeliverablesDocument");
          const uploadResult = await folder.files.addUsingPath(
            `${Date.now()}_${fileObj.file.name}`,
            fileObj.file,
            { Overwrite: true }
          );

          // Get list item associated with uploaded file
          // Get list item associated with uploaded file
          const fileItem = await uploadResult.file.getItem();
          const uploadedFileItemId = (fileItem as any).Id;

          // Update file metadata
          await spContext.web.lists.getByTitle("DeliverablesDocument").items
            .getById(uploadedFileItemId)
            .update({
              ProjectID: selectedTask.projectId.toString(),
              DeliverablesDetailsId: selectedTask.deliverableId.toString(),
              Revision: newRevision,

            });

          uploadedFileIds.push(uploadedFileItemId);

          console.log("File metadata updated successfully");
        }
      }
      if (existingDocuments.length > 0) {
        for (const doc of existingDocuments) {
          await spContext.web.lists.getByTitle("DeliverablesDocument").items
            .getById(doc.id)
            .update({

            });
        }
      }
      // STEP 2: Update DeliverablesDetails item with new revision
      const primaryFileId = uploadedFileIds.length > 0 ? uploadedFileIds[0] : null;
      await spContext.web.lists.getByTitle("DeliverablesDetails").items
        .getById(selectedTask.deliverableId)
        .update({
          DocumentComments: comment,
          RevisionNumber: newRevision,
          Status: "In-Progress",
          IsReworked: "",
          DeliverablesDocumentIDId: uploadedFileIds // Keep for backward compatibility
        });

      console.log("DeliverablesDetails updated successfully");

      // STEP 3: Mark Vendor's approval as Completed
      const projectId = selectedTask?.projectId;
      const deliverableId = selectedTask?.deliverableId;
      const approverType = "Vendor";
      const currentUser = await sp.web.currentUser();

      try {
        const items = await spContext.web.lists
          .getByTitle("ProjectApprovals")
          .items.select(
            "*",
            "DeliverablesDetailsId/ID",
            "ProjectCreationListID/ID",
            "AssignedTo/ID",
            "AssignedTo/Title",
            "AssignedTo/EMail"
          )
          .expand("DeliverablesDetailsId", "ProjectCreationListID", "AssignedTo")
          .orderBy("SerialNumber", true)
          .filter(
            `ProjectCreationListID/ID eq ${projectId} and ` +
            `DeliverablesDetailsId/ID eq ${deliverableId} and ` +
            `AssignedTo/ID eq ${currentUser.Id} and ` +
            `ApproverRole eq '${approverType}' and ` +
            `Status eq 'Pending'`
          )();

        if (items.length > 0) {
          const vendorTaskId = items[0].Id;
          await spContext.web.lists.getByTitle("ProjectApprovals").items.getById(vendorTaskId).update({
            Status: "Completed",
            Remarks: comment,
            ApprovalDate: new Date()
          });
          console.log("Vendor task marked as completed");
        } else {
          console.log("No matching ProjectApprovals found for current user");
        }
      } catch (error) {
        console.error("Error updating ProjectApprovals:", error);
      }
      // vishnu added
      const projectItem = await spContext.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(selectedTask.projectId)
        .select("PreparedBy/Id", "PreparedBy/Title", "PreparedBy/EMail")
        .expand("PreparedBy")();

      const preparedByUserId = projectItem.PreparedBy?.Id;
      const preparedByUserName = projectItem.PreparedBy?.Title;

      // vishnu ended
      // STEP 4: Create new ProjectApproval item for Document Controller with new revision
      await spContext.web.lists.getByTitle("ProjectApprovals").items.add({
        // 🔹 Get Prepared By user (from ProjectCreationList)



        DeliverablesDetailsIdId: selectedTask.deliverableId,
        ProjectCreationListIDId: selectedTask.projectId,
        AssignedToId: documentControllerId,
        DocumentType: selectedTask.docType,
        ApproverRole: "Document Controller",
        ProjectType: selectedTask.projectType,
        Level: "Level 1",
        SerialNumber: 0,
        ApprovalCriteria: "Anyone",
        RequestedById: currentUser.Id,
        RequestedDate: new Date(),
        IncomingDate: new Date(),
        RequestedRole: "Vendor",
        RevisionNumber: newRevision,
        Status: "Pending"
      });

      // satish added
      const fileDetails: { name: string; link: string, docrevisionNo: string }[] = [];
      for (const id of uploadedFileIds) {
        const details = await waitForShareLink(id);
        // Clean the file name by removing the prefix before first underscore
        const originalName = details.fileName;
        const cleanedName = originalName.includes("_")
          ? originalName.substring(originalName.indexOf("_") + 1)
          : originalName;
        fileDetails.push({
          name: cleanedName,
          link: details.shareLink,
          docrevisionNo: details.revisionNo
        });
      }

      const actionType: ActionType = "Submitted";

      const emailBody = buildApprovalEmailBody(
        actionType,
        documentControllerName, // whoever is receiving
        preparedByUserName,   // whoever is sending
        fileDetails.map(d => ({
          deliverable: selectedTask?.deliverable,
          fileName: d.name,
          sharedLink: d.link,
          docrevisionNo: d.docrevisionNo
        }))
      );
      // satish added ended


      // Vishnu Added
      await addEmailTriggerDetails(
        selectedTask.projectName,
        documentControllerId!,
        preparedByUserId,
        emailBody,
        uploadedFileIds
      );
      // vishnu ended

      setShowSubmitLoader(false); // ✅ ADD
      setPopup({
        isOpen: true,
        type: 'success',
        title: 'Success',
        message: 'Task submitted successfully.',
        onConfirm: () => {
          setPopup(prev => ({ ...prev, isOpen: false }));
          handleBackClick();
          window.location.reload();
        }
      });

    } catch (error) {
      console.error("Error in submission:", error);
      setShowSubmitLoader(false); // ✅ STOP loader
      setPopup({
        isOpen: true,
        type: 'error',
        title: 'Error',
        message: 'An error occurred during submission. Please check console for details.'
      });
    }
  }

  // satish added

  // const buildApprovalEmailBody = (
  //   actionType: ActionType,
  //   approverName: string,
  //   senderName: string,
  //   docs: { deliverable: string; fileName: string; sharedLink: string; }[]
  // ) => {
  //   const { message } = actionConfig[actionType];

  //   const rows = docs.map(d => `
  //     <tr>
  //       <td style="border:1px solid #ccc;padding:8px;">${d.deliverable}</td>
  //       <td style="border:1px solid #ccc;padding:8px;">
  //         <a href="${d.sharedLink}" target="_blank">${d.fileName}</a>
  //       </td>
  //     </tr>
  //   `).join("");

  //   return `
  //     Dear ${approverName},<br/><br/>
  //     ${message}<br/><br/>

  //     <table style="border-collapse:collapse;width:100%;max-width:700px;">
  //       <thead>
  //         <tr>
  //           <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Deliverable</th>
  //           <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Document</th>
  //         </tr>
  //       </thead>
  //       <tbody>
  //         ${rows}
  //       </tbody>
  //     </table>

  //     <br/><br/>

  //     <a href="https://officeindia.sharepoint.com/sites/multiverseintranetportal/SitePages/MyApprovals.aspx" target="_blank">
  //       Click here to view in portal
  //     </a>

  //     <br/><br/>
  //     Regards,<br/>
  //     ${senderName}
  //   `;
  // };
  const buildApprovalEmailBody = (
    actionType: ActionType,
    approverName: string,
    senderName: string,
    docs: { deliverable: string; fileName: string; sharedLink: string; docrevisionNo: string }[]
  ) => {
    const { message } = actionConfig[actionType];

    // Format current date as dd MMM yyyy
    const today = new Date();
    const sentDate = today.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric"
    });

    let rows = docs.map((d, i) => {
      const no = i + 1;
      const transmittal = `TR-${String(no).padStart(3, "0")}`;
      const remarks = "Submitted for Review";

      // <td style="border:1px solid #ccc;padding:6px;">${selectedTask?.docNumber}</td>
      return `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td>
          <td style="border:1px solid #ccc;padding:6px;">
            <a href="${d.sharedLink}" target="_blank">${selectedTask?.docNumber}</a>
          </td>
          <td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${d.docrevisionNo ?? "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
          <td style="border:1px solid #ccc;padding:6px;">${transmittal}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">Attached</td>
        </tr>
      `;
    }).join("");

    // --- Extra Dropbox Row ---
    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;
    const dropboxUrl = `https://officeindia.sharepoint.com/sites/multiverseintranetportal/DeliverablesDoc/Forms/AllItems.aspx?id=%2Fsites%2Fmultiverseintranetportal%2FDeliverablesDoc%2FESSADeliverablesDoc&viewid=21fa2859%2Dc815%2D46d1%2D96bd%2Dbdabee445848&npsAction=createList`;

    rows += `
    <tr>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td>
      <td style="border:1px solid #ccc;padding:6px;"></td>
      <td style="border:1px solid #ccc;padding:6px;">${selectedTask.projectName}</td>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
      <td style="border:1px solid #ccc;padding:6px;">${extraTransmittal}</td>
      <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
      <td style="border:1px solid #ccc;padding:6px;">
        <a href="${dropboxUrl}" target="_blank">Dropbox Link</a>
      </td>
    </tr>
  `;

    return `
      Dear ${approverName},<br/><br/>
      ${message}<br/><br/>
  
      <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;">
        <thead>
          <tr style="background:#8eaada;font-weight:bold;">
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
            <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
            <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
            <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
            <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
            <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>
  
      <br/><br/>
  
      <a href="https://officeindia.sharepoint.com/sites/multiverseintranetportal/SitePages/MyApprovals.aspx" target="_blank">
        Click here to view in portal
      </a><br/><br/>
  
      Regards,<br/>
      ${senderName}
    `;
  };


  const waitForShareLink = async (fileItemId: number, maxAttempts = 20, delayMs = 3000) => {
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const item = await spContext.web.lists
        .getByTitle("DeliverablesDocument")
        .items.getById(fileItemId)
        .select("File/Name", "SharedLink", "Revision")
        .expand("File")();

      if (item.SharedLink) {
        return {
          shareLink: item.SharedLink,
          fileName: item.File?.Name,
          revisionNo: item.Revision
        };
      }

      // wait before next attempt
      await new Promise(res => setTimeout(res, delayMs));
    }

    throw new Error(`ShareLink not generated for FileItemId: ${fileItemId}`);
  };

  // satish addded end

  // vishu Added
  // 🔹 EMAIL TRIGGER DETAILS (My Task)
  const addEmailTriggerDetails = async (
    projectName: string,
    documentControllerId: number,
    preparedByUserId: number,
    emailbody: string,
    attachmentIds: number[]
  ) => {
    const currentUser = await spContext.web.currentUser();

    await spContext.web.lists
      .getByTitle("EmailTriggerDetails")
      .items.add({
        Title: projectName,
        // Body: `The assigned task has been completed and the document has been submitted for your review. Kindly review the submitted document and proceed with the next steps.`,
        Body: emailbody,
        Subject: `Document Submission Confirmation – ${selectedTask?.deliverable}`,

        // ✅ ToUser = Document Controller
        ToUserId: [documentControllerId],

        // ✅ CCUser = Logged-in user + Prepared By
        CCUserId: [currentUser.Id, preparedByUserId],

        // AttachmentIdId: attachmentIds

      });
  };
  // vishnu ended


  return (
    <div id="wrapper" ref={elementRef}>
      {/* 🔄 SUBMIT TASK LOADER */}
      {showSubmitLoader && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0,0,0,0.45)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999,
          }}
        >
          <div className="text-center">
            <img
              src={loaderGif}
              alt="Submitting..."
              style={{ width: "90px", height: "90px" }}
            />
            <p className="text-white mt-3 fw-semibold">
              Submitting, please wait...
            </p>
          </div>
        </div>
      )}

      <div style={{ display: 'none' }} className="app-menu" id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>

      <div className="content-page">
        <HorizontalNavbar _context={sp} className="app-menu-hide" siteUrl={siteUrl} />
        <div
          className="content"
          style={{
            marginLeft: `${!useHide ? "15px" : "0px"}`,

            marginTop: "0rem",
          }}
        >
          <div className="container-fluid  paddb">
            <img style={{
              position: 'fixed', top: '8px', left: '1px', zIndex: '99'
            }}
              src={newlogo} />
            <div className={styles.myTask}>
              {!showForm ? (
                <>
                  <div className='row mb-2'>
                    <div className='col-md-6'>
                      <h4 className='page-title fw-bold mb-1 text-dark font-20'>Vendor Portal</h4>
                      <ol className="breadcrumb mb-0">
                        {/* <li className="breadcrumb-item">
                          <a href="">Home</a> </li>
                        <li className="breadcrumb-item active"><a href="#">My Task</a></li> */}
                      </ol>

                    </div>



                  </div>

                  <div className='row nrefield'>
                    <div className='col-sm-4'>
                      <div className="card">
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>

                            <p style={{ fontSize: '20px' }} className="mb-2 text-dark font-20 fw-bold">Total Documents</p>
                            <h3 style={{ fontSize: '16px' }} className="mb-0 font-16 text-dark"><span data-target="438">{TotalDocumentCount}</span></h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={totalim} alt='projects' />
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className='col-sm-4'>
                      <div style={{ width: '100%' }}
                        className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ''}`}
                        onClick={() => getDeliverablesDetails("Approved")}
                      >
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>

                            <p style={{ fontSize: '20px' }} className="mb-2 text-dark font-20 fw-bold">Completed</p>
                            <h3 style={{ fontSize: '16px' }} className="mb-0 font-16 text-dark"><span data-target="438">{completedCount}</span></h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={approve} alt='projects' />
                          </div>
                        </div>
                      </div>
                    </div>


                    {/* <div className='col-sm-4'>
                      <div className="card">
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>

                            <p style={{ fontSize: '20px' }} className="mb-2 text-dark font-20 fw-bold">Pending</p>
                            <h3 style={{ fontSize: '16px' }} className="mb-0 font-16 text-dark"><span data-target="438">48</span></h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={pen} alt='projects' />
                          </div>
                        </div>
                      </div>
                    </div> */}
                    <div className='col-sm-4'>
                      <div style={{ width: '100%' }}
                        className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ''}`}
                        onClick={() => getDeliverablesDetails("Pending")}
                      >
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>
                            <p style={{ fontSize: '20px' }} className="mb-2 text-dark font-20 fw-bold">Pending</p>
                            <h3 style={{ fontSize: '16px' }} className="mb-0 font-16 text-dark">
                              {pendingCount}
                            </h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={pen} alt='projects' />
                          </div>
                        </div>
                      </div>
                    </div>


                  </div>

                  <div className='card mt-0'>
                    <div className='card-body pb-0'>
                      <div className='d-flex align-items-center justify-content-between'>
                        <div className={styles.tilesContainer}>
                          <div
                            className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ''}`}
                            onClick={() => getDeliverablesDetails("Pending")}
                          >
                            <div className={styles.tileBody}>
                              <div className={styles.tileRow}>
                                <div className={styles.tileContent}>
                                  <p className={styles.tileLabel}>Pending</p> <h3 className={styles.tileCount}>{pendingCount}</h3>
                                </div>
                              </div>
                            </div>
                          </div>

                          <div
                            className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ''}`}
                            onClick={() => getDeliverablesDetails("Approved")}
                          >
                            <div className={styles.tileBody}>
                              <div className={styles.tileRow}>
                                <div className={styles.tileContent}>

                                  <p className={styles.tileLabel}>Completed</p>
                                  <h3 className={styles.tileCount}>{completedCount}</h3>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div style={{ marginTop: '-12px' }} className='d-flex align-items-center justify-content-end gap-2'>

                          <div>
                            <div className='form-group d-flex align-items-center justify-content-end gap-1'>
                              <label style={{ width: '100px' }} className='mb-1'>From Date</label>
                              <input type='date' className='form-control' />
                            </div>
                          </div>
                          <div className='form-group d-flex align-items-center justify-content-end gap-1'>
                            <label style={{ width: '100px' }} className='mb-1'>Due Date</label>
                            <input type='date' className='form-control' />
                          </div>
                          <div className='btn btn-primary pr7'>
                            Filter
                          </div>

                        </div>
                      </div>
                    </div></div>
                  {/* Tiles */}


                  {/* Table */}
                  <div style={{ clear: 'both', marginTop: '15px' }} className='mt-2'>

                    <div className=''>
                      <div className={styles.mainTableContainer}>
                        <div className={styles.tableCard}>
                          <div className={styles.tableWrapper}>
                            <table className={styles.taskTable}>
                              <thead>
                                <tr>
                                  <th style={{ minWidth: '70px', textAlign: 'center' }}>S.No</th>
                                  <th style={{ textAlign: 'center' }}>Project Name</th>
                                  <th>Project Type</th>
                                  <th>Deliverable</th>
                                  <th style={{ textAlign: 'center' }}>Document Type</th>
                                  <th style={{ textAlign: 'center' }}>Due Date</th>
                                  <th>Organization</th>
                                  <th style={{ minWidth: '90px', textAlign: 'center' }}>Status</th>
                                  <th style={{ minWidth: '70px', textAlign: 'center' }}>Action</th>
                                </tr>
                              </thead>
                              <tbody>
                                {filteredTasks.map((task, index) => (
                                  <tr key={task.sno}>
                                    <td style={{ minWidth: '70px' }}>
                                      <div className="indexdesign" style={{ marginLeft: '8px' }}>{index + 1} </div></td>
                                    <td style={{ textAlign: 'center' }}>{task.projectName}</td>
                                    <td>{task.projectType}</td>
                                    <td>{task.deliverable}</td>
                                    <td style={{ textAlign: 'center' }}>{task.docType}</td>
                                    <td>
                                      {task.DueDate
                                        ? new Date(task.DueDate).toLocaleDateString('en-GB')
                                        : '-'}
                                    </td>

                                    <td>{task.org}</td>
                                    <td style={{ minWidth: '90px', textAlign: 'center' }}>
                                      <span className={`${styles.statusBadge} ${getStatusClass(task.status)}`}>
                                        {task.status}
                                      </span>
                                    </td>
                                    <td style={{ minWidth: '70px', textAlign: 'center' }}>
                                      <img
                                        src={eye} style={{ cursor: 'pointer' }}
                                        onClick={() => handleViewClick(task)}
                                        title={`View ${task.deliverable}`}
                                      />
                                      {/* <span
                                        className={styles.actionIcon}
                                       
                                      >
                                        👁️
                                      </span> */}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                        {/* <div className={styles.scrollHint}><em>Scroll horizontally → to view all columns</em></div> */}
                      </div>
                    </div>
                  </div>

                </>
              ) : (
                <>
                  {/* Form View */}
                  <div className={styles.formContainer}>
                    <div className={styles.formHeader}>
                      <h3 style={{ margin: '0px' }}>
                        My Task &gt;&gt; {selectedTask?.docNumber}
                      </h3>
                      <button type="button" className={styles.backButton} onClick={handleBackClick}>Back</button>
                    </div>
                    <div className={styles.formView}>
                      <div className={styles.formInner}>
                        <div className={styles.formBody}>
                          <div className={styles.formGrid}>
                            <div>
                              <label>Project Name</label>
                              <input type="text" value={selectedTask?.projectName || ''} disabled />
                            </div>
                            <div>
                              <label>Project Type</label>
                              <input type="text" value={selectedTask?.projectType || ''} disabled />
                            </div>
                            <div>
                              <label>Client Name</label>
                              <input type="text" value={selectedTask?.clientName || ''} disabled />
                            </div>
                            <div>
                              <label>Project Date</label>
                              <input type="text" value={selectedTask?.creationDate ? new Date(selectedTask.creationDate).toLocaleDateString('en-GB') : ''} disabled />
                            </div>
                            <div>
                              <label>Prepared By</label>
                              <input
                                type="text"
                                value={selectedTask?.preparedBy || ''}
                                disabled

                              />
                            </div>
                            <div>
                              <label>Deliverable</label>
                              <input type="text" value={selectedTask?.deliverable || ''} disabled />
                            </div>
                            <div>
                              <label>Document Type</label>
                              <input type="text" value={selectedTask?.docType || ''} disabled />
                            </div>
                            <div>
                              <label>Document Number</label>
                              <input
                                type="text"
                                value={selectedTask?.docNumber || ''}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Area</label>
                              <input type="text" value={selectedTask?.area || ''} disabled />
                            </div>
                            <div>
                              <label>Organization</label>
                              <input type="text" value={selectedTask?.org || ''} disabled />
                            </div>
                            <div>
                              <label>Revision Number</label>
                              <input
                                type="text"
                                value={
                                  selectedTask?.status === "Pending" && existingDocuments
                                    ? (parseInt(selectedTask?.revisionnumber || "0") + 1).toString()
                                    : selectedTask?.revisionnumber || "0"
                                }
                                disabled
                              />
                            </div>
                          </div>

                          {/* Upload & Comment Section में बदलाव */}
                          <div className={styles.formActions}>
                            {/* Existing Documents Section - for Pending tasks */}
                            {selectedTask?.status === "Pending" && existingDocuments.length > 0 && (
                              <>
                                <label>Previously Uploaded Documents:</label>
                                <div className={styles.existingFilesList}>
                                  {existingDocuments.map((doc, index) => (
                                    <p key={index}>
                                      <a href={doc.url} target="_blank" rel="noopener noreferrer">
                                        {doc.name}
                                      </a>
                                    </p>
                                  ))}
                                </div>
                              </>
                            )}

                            {/* For Completed/In-Progress tasks - show all uploaded files */}
                            {(selectedTask?.status === "Approved" || selectedTask?.status === "In-Progress") && existingDocuments.length > 0 && (
                              <>
                                <label>All Uploaded Documents:</label>
                                <div className={styles.existingFilesList}>
                                  {existingDocuments.map((doc, index) => (
                                    <p key={index}>
                                      <a href={doc.url} target="_blank" rel="noopener noreferrer">
                                        {doc.name}
                                      </a>
                                    </p>
                                  ))}
                                </div>
                              </>
                            )}

                            {/* Show upload field only for Pending tasks */}
                            {selectedTask?.status === "Pending" && (
                              <>
                                <label>Upload Documents (Multiple)*</label>
                                <input
                                  type="file"
                                  multiple
                                  onChange={(e) => {
                                    if (e.target.files) {
                                      const filesArray = Array.from(e.target.files).map(file => ({
                                        name: file.name,
                                        file: file,
                                        url: undefined as string | undefined
                                      }));
                                      setSelectedFiles(prev => [...prev, ...filesArray]);
                                    }
                                  }}
                                />

                                {/* Show selected files list */}
                                {selectedFiles.length > 0 && (
                                  <div className={styles.selectedFilesList}>
                                    <label>Selected Files:</label>
                                    {selectedFiles.map((file, index) => (
                                      <div key={index} className={styles.fileItem}>
                                        <span>{file.name}</span>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            const newFiles = [...selectedFiles];
                                            newFiles.splice(index, 1);
                                            setSelectedFiles(newFiles);
                                          }}
                                          className={styles.removeFileBtn}
                                        >
                                          ✕
                                        </button>
                                      </div>
                                    ))}
                                  </div>
                                )}
                              </>
                            )}

                            <label>Comment*</label>
                            <textarea
                              placeholder="Enter your comment"
                              value={comment}
                              onChange={(e) => setComment(e.target.value)}
                              disabled={selectedTask?.status === "Approved" || selectedTask?.status === "In-Progress"}
                            />
                          </div>

                          {/* Document Comments Accordion */}
                          {showDocumentComments && (
                            <div className={styles.accordionItem}>
                              <h2 className={styles.accordionHeader}>
                                <div className={styles.accordionButton}>
                                  Document Comments
                                </div>
                              </h2>

                              <div className={styles.accordionBody}>
                                <div className={styles.customCard}>
                                  <div className={styles.documentCommentsHeader}>
                                    <select
                                      className={styles.formSelect}
                                      value={selectedVersion}
                                      onChange={(e) => onVersionChange(e.target.value)}
                                    >
                                      <option value="">-- Select Revision --</option>
                                      {versionList.map(version => (
                                        <option key={version} value={version}>{version}</option>
                                      ))}
                                    </select>
                                    <button
                                      className={styles.btnOutlineSuccess}
                                      type="button"
                                      onClick={exportCommentsToExcel}
                                    >
                                      Export to Excel
                                    </button>
                                    <button
                                      className={styles.btnOutlineSuccess}
                                      type="button"
                                      onClick={refreshDocComment}
                                    >
                                      ↻
                                    </button>
                                  </div>

                                  <div className={styles.ribbonContent}>
                                    <table className={styles.commentsTable}>
                                      <thead>
                                        <tr>
                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Users</th>
                                          <th style={{ minWidth: '100px', maxWidth: '100px' }}>Comment Date</th>
                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Page No.</th>
                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Revision</th>
                                          <th style={{ minWidth: '200px', maxWidth: '200px' }}>Comments</th>
                                        </tr>
                                      </thead>
                                      <tbody style={{ maxHeight: '250px', overflowY: 'auto' }}>
                                        {documentComments.map((commentItem) => (
                                          <tr key={commentItem.id}>
                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                              {commentItem.userName}
                                            </td>
                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '100px', maxWidth: '100px' }}>
                                              {commentItem.commentDate}
                                            </td>
                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                              {commentItem.pageNumber}
                                            </td>
                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                              {commentItem.revision}
                                            </td>
                                            <td style={{ padding: '15px', verticalAlign: 'top', minWidth: '200px', maxWidth: '200px' }}>
                                              {commentItem.comment}
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

                          {/* Audit History */}
                          <div className={styles.auditHistory}>
                            <h4>Audit History</h4>
                            {showNoAuditHistory && auditHistory.length === 0 ? (
                              <p>No audit history available</p>
                            ) : (
                              <table className='mtablemyt'>
                                <thead>
                                  <tr>
                                    <th style={{ minWidth: '70px', maxWidth: '70px' }}>SNo</th>
                                    <th>Approval Level</th>
                                    <th>Assigned To</th>
                                    <th>Assigned To Role</th>
                                    <th>Requestor Name</th>
                                    <th>Requested Date</th>
                                    <th>Action Taken By</th>
                                    <th>Action Taken On</th>
                                    <th>Remark</th>
                                    <th style={{ minWidth: '110px', maxWidth: '110px' }}>Status</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {auditHistory.map((historyItem) => (
                                    <tr key={historyItem.sno}>
                                      <td style={{ minWidth: '70px', maxWidth: '70px' }}>{historyItem.sno}</td>
                                      <td>{historyItem.approvalLevel}</td>
                                      <td>{historyItem.assignedTo}</td>
                                      <td>{historyItem.assignedToRole}</td>
                                      <td>{historyItem.requestorName}</td>
                                      <td>{historyItem.requestedDate}</td>
                                      <td>{historyItem.assignedTo}</td>
                                      <td>{historyItem.actionTakenOn}</td>
                                      <td>{historyItem.remark}</td>
                                      <td style={{ minWidth: '110px', maxWidth: '110px' }}>
                                        <span className={`${styles.statusBadge} ${getStatusClass(historyItem.status)}`}>
                                          {historyItem.status}
                                        </span>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            )}
                          </div>

                          <div className={styles.formButtons}>
                            {selectedTask?.status === "Pending" && (
                              <>
                                <button
                                  type="button"
                                  className={styles.submitButton}
                                  onClick={handleSubmitClick}
                                  disabled={showSubmitLoader}
                                >Submit</button>

                                <button type="button" className={styles.cancelButton} onClick={handleBackClick}>Cancel</button>
                              </>
                            )}
                            {selectedTask?.status === "Approved" && (
                              <button type="button" className={styles.cancelButton} onClick={handleBackClick}>Back</button>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}
              <CustomPopup
                isOpen={popup.isOpen}
                type={popup.type}
                title={popup.title}
                message={popup.message}
                onConfirm={popup.onConfirm}
                onCancel={() => setPopup(prev => ({ ...prev, isOpen: false }))}
                onClose={() => setPopup(prev => ({ ...prev, isOpen: false }))}
                onSuccessOk={() => {
                  setPopup(prev => ({ ...prev, isOpen: false }));
                  if (popup.onConfirm) popup.onConfirm();
                }}
              />
            </div>
          </div></div>
      </div>
    </div>
  );
};

export default MyTask;