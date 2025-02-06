import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { getSP } from '../webparts/addBusinessApp/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
const sp = getSP();
let web = "";
// import { sp } from "@pnp/sp/presets/all";  // Import PnPjs
// import { SPFetchClient } from "@pnp/sp/authenticators";
// sp.configure({
//   sp: {
//     baseUrl: siteUrl,
//   },
// });
export const fetchAutomationDepartment = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGAutomationDepartment").items.getAll().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const addItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const getCategory = async (_sp, id) => {
  let arr = []
  await _sp.web.lists.getByTitle("BusinessAppsCategory").items.select("ID,CategoryName").expand("").filter(`(Active eq 1) and(AnnouncementandNewsTypeMaster/ID eq ${id})`)()
    .then((res) => {
      console.log(res);
      const newArray = res.map(({ ID, CategoryName }) => ({ id: ID, name: CategoryName }));
      console.log(newArray, 'newArray');

      arr = newArray;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchAutomationCategory = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("BusinessAppsCategory").items.getAll().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchARGAutomationdata = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGBusinessApps").items.select("*,Category/Id,Category/CategoryName")
    .expand("Category")
    .orderBy("Order0", true)
    .getAll().then((res) => {
      console.log("response-->>>", res);

      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

// export const getMyApprovalsdata = async (_sp,listName,status) => {
//   let arr = []
//   let currentUser;
//   await _sp.web.currentUser()
//     .then(user => {
//       console.log("user",user);
//       currentUser = user.Email; 
//       // Get the current user's Email
//     })
//     .catch(error => {
//       console.error("Error fetching current user: ", error);
//       return [];
//     });

//   if (!currentUser) return arr; 
//   // Return empty array if user fetch failed

//   await _sp.web.lists.getByTitle(listName).items
//     .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
//     .filter(`AssignedTo/EMail eq '${currentUser}' and Status eq '${status}'`)      
//     .orderBy("Created", false).getAll()
//     .then((res) => {
//       console.log(`--MyApproval${listName}`, res);
//       arr = res
//       // arr = res.filter(item => 
//       //     // Include public groups or private groups where the current user is in the InviteMembers array
//       //     item.GroupType === "Public" || 
//       //     (item.GroupType === "Private" && item.InviteMemebers && item.InviteMemebers.some(member => member.Id === currentUser))
//       //   );
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }
//Add Business Apps


// Initialize PnPjs with the current site collection


// Function to fetch data from a single SharePoint site collection
export const getListDataFromSiteCollection = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  // Setup PnPJs for a specific site collection URL
  let arr = [];
  console.log("sdsssss", sp, _sp)
  debugger
  let apiUrl;
  let FinalStatus = "";
  if (status == "Pending") {
    FinalStatus = "Not Started"
  } else if (status == "Approved") {
    FinalStatus = "Completed"
  }
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${Actingfor}' and TaskStatus eq '${FinalStatus}'`}`;
  } else {
    let currentUser;
    await _sp.web.currentUser()
      .then(user => {
        console.log("user", user);
        currentUser = user.Email; // Get the current user's Email
      })
      .catch(error => {
        console.error("Error fetching current user: ", error);
        return [];
      });

    if (!currentUser) return arr; // Return empty array if user fetch failed
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${currentUser}' and TaskStatus eq '${FinalStatus}'`}`;
  }
  try {
    console.log("apiUrlapiUrl", apiUrl);
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    // Wait for the response to be converted to JSON
    const data = await response.json();

    // Immediately handle data
    console.log('List Items ellllllllnew:', data.value);

    // You can now manipulate or return the data as needed
    if (data.value.length > 0) {
      arr = data.value;
    } else {
      arr = []
    }
    return arr;  // Directly returning the array if needed
  } catch (error) {
    console.error('Error fetching data:', error);
  }
}

export const getDataFromMultipleSites = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  const allData = [];
  const siteUrls = [
    SiteBaseURL
  ];
  // Loop through each site collection URL and fetch data
  for (const siteUrl of siteUrls) {
    const data = await getListDataFromSiteCollection(_sp, listName, status, Actingfor, portal, siteUrl);
    allData.push(...data);
    console.log("dadadadadad", data);
  }

  console.log('Combined data from all site collections:', allData);
  return allData;
}

// Call the function to get the data


export const getMyApprovalsdata = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  let arr = []
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    await _sp.web.lists.getByTitle(listName).items
      .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
      .filter(`AssignedTo/EMail eq '${Actingfor}' and Status eq '${status}'`)
      .orderBy("Created", false).getAll()
      .then((res) => {
        console.log(`--MyApproval${listName}`, res);
        arr = res
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });

    return arr;


  } else {
    //let arr = []
    let currentUser;
    await _sp.web.currentUser()
      .then(user => {
        console.log("user", user);
        currentUser = user.Email; // Get the current user's Email
      })
      .catch(error => {
        console.error("Error fetching current user: ", error);
        return [];
      });

    if (!currentUser) return arr; // Return empty array if user fetch failed

    await _sp.web.lists.getByTitle(listName).items
      .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
      .filter(`AssignedTo/EMail eq '${currentUser}' and Status eq '${status}'`)
      .orderBy("Created", false).getAll()
      .then((res) => {
        console.log(`--MyApproval${listName}`, res);
        arr = res
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;

  }
  console.log("arrrrrrr", arr)
}

export const updateItem = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};

export const uploadFileToLibrary = async (file, sp, docLib) => {

  let arrFIleData = [];
  let fileSize = 0
  try {
    const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
        fileSize = progress.fileSize
      },
      true
    );

    const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*", "ID", "AuthorId", "Modified")
    console.log(item.Id, 'itemitem');
    let arr = {
      ID: item.Id,
      Createdby: item.AuthorId,
      Modified: item.Modified,
      fileUrl: result.data.ServerRelativeUrl,
      fileSize: fileSize,
      fileType: file.type,
      fileName: file.name,
    }
    arrFIleData.push(arr)
    console.log(arrFIleData);

    return arrFIleData;
  } catch (error) {
    console.log("Error uploading file:", error);
    return null; // Or handle error differently
  }
};
export const uploadFile = async (file, sp, docLib, siteUrl) => {
  let arr = {};

  const uploadResult = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true);

  const fileUrl = uploadResult.data.ServerRelativeUrl;

  const imgMetadata = {
    "__metadata": { "type": "SP.FieldUrlValue" },
    "Description": file.name,
    "Url": `${siteUrl}${fileUrl}`
  };

  // await sp.web.lists.getByTitle(docLib).items.getById(uploadResult.data.UniqueId).update({
  //   "AnnouncementandNewsBannerImage": imgMetadata
  // });
  arr = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "Image",
    "serverRelativeUrl": fileUrl
  };
  return arr;
};
export const getBusinessAppsByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  await _sp.web.lists.getByTitle("ARGMediaGallery").items.getById(id).select("*,EntityMaster/Id,EntityMaster/Entity,MediaGalleryCategory/Id,MediaGalleryCategory/CategoryName").expand("EntityMaster,MediaGalleryCategory")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      const parsedValues = {
        Title: res.Title,
        ID: res.ID,
        entity: res.EntityMaster?.Id,
        Image: res.Image,
        MediaGalleriesId: res?.MediaGalleriesId,
        MediaGalleryJSON: res?.MediaGalleryJSON,
        Category: res?.MediaGalleryCategory?.Id
        // other fields as needed
      };
      arr.push(parsedValues)
      arrs = arr
      console.log(arrs, 'arr');
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arrs, 'arr');
  return arrs;
}
export const getUrl = async (sp, siteUrl) => {
  let srt = siteUrl;
  let dynamicPart = "/sites/";
  let UrlArr = [];
  try {
    let index = srt.lastIndexOf(dynamicPart); // Find the last occurrence of "/sites/"

    if (index !== -1) {
      let endIndex = srt.indexOf("/", index + dynamicPart.length) !== -1
        ? srt.indexOf("/", index + dynamicPart.length)
        : srt.length;

      let updatedStr = srt.slice(0, index) + srt.slice(endIndex);
      console.log(updatedStr, 'updatedStr');
      const url = await sp.web.currentUser.getContextInfo();
      console.log(url, 'res');

      let UrlArr1 =
      {
        DomainUrl: updatedStr,
        WebFullUrl: url.WebFullUrl

      }
      UrlArr.push(UrlArr1)
    } else {
      console.log("Pattern not found. No replacement was made.");
    }
  } catch (error) {
    console.log("An error occurred:", error.message);
  }

  return UrlArr
}
export const ARGBusinessAppCategory = async (sp) => {

  let arr = []
  let arrs = []
  await sp.web.lists.getByTitle("BusinessAppsCategory").items.orderBy("Created", false).getAll().then((res) => {
    console.log(res, 'Resss');

    arrs = res
    console.log(arrs, 'arr');
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arrs, 'arr');
  return arrs;
}

//End
//Business apps Master
export const getBusinessApps = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGBusinessApps")
    .items.select("*,EntityMaster/ID,EntityMaster/Entity,Category/ID,Category/CategoryName").expand("EntityMaster,Category").orderBy("Created", false).getAll()
    .then((res) => {
      console.log(res);

      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const DeleteBusinessAppsAPI = async (_sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.getById(id).delete();
    console.log('Item deleted successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
}
//End
// export const getApprovalListsData = async (_sp,status) => {
//   let arr = []

//   await _sp.web.lists.getByTitle("AllApprovalLists").items.orderBy("Created", false).getAll()
//     .then(async (res) => {
//       console.log("AllApprovallists",res);
//       let AllApprovalArr = [];

//       for (let i = 0; i < res.length; i++) {
//        await getMyApprovalsdata(_sp,res[i].Title,status).then((resData)=>{
//           for (let j = 0; j < resData.length; j++) {
//             AllApprovalArr.push(resData[j])
//           }

//         })
//       }
//       console.log("AllApprovalArr",AllApprovalArr);
//       arr = AllApprovalArr;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }
export const getApprovalListsData = async (_sp, status, Actingfor) => {
  let arr = []
  debugger

  if (!Actingfor) {
    // alert(`acting for ${Actingfor} is not null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];

        for (let i = 0; i < res.length; i++) {

          if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
            await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
              .then((resData) => {
                if (resData && resData.length > 0) {
                  for (let j = 0; j < resData.length; j++) {
                    // AllApprovalArr.push(resData[j])
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].Title,
                      ApprovalTitle: "",
                      Author: resData[j].Requestor_x0020_Name,
                      ProcessName: res[i].ProcessName,
                      Created: resData[j].Created,
                      Status: resData[j].TaskStatus,
                      TaskID: resData[j].ID,
                      AppID: res[i].AppId,
                      RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                    })
                    // AllApprovalArr.push({
                    //   ID: resData[j].ID,
                    //   RequestID: resData[j].ID,
                    //   ApprovalTitle: resData[j].Title,
                    //   Author: resData[j].Requestor_x0020_Name,
                    //   ProcessName: res[i].ProcessName,
                    //   Created: resData[j].Created,
                    //   Status: resData[j].TaskStatus,
                    //   TaskID: resData[j].MasterID,
                    //   AppID: res[i].AppId,
                    //   RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                    // })
                  }
                }
                // Handle the combined data here
                console.log("Final combined data:", resData);
              })
              .catch((error) => {
                console.error("Error fetching data:", error);
              });


          } else {
            await getMyApprovalsdata(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
              if (resData && resData.length > 0) {
                console.log("resDataresDataresDataresData", resData);
                for (let j = 0; j < resData.length; j++) {
                  AllApprovalArr.push({
                    ID: resData[j].ID,
                    RequestID: resData[j].RequestID,
                    ApprovalTitle: resData[j].ApprovalTitle,
                    Author: resData[j].Author,
                    ProcessName: resData[j].ProcessName,
                    Created: resData[j].Created,
                    Status: resData[j].Status,
                    TaskID: "",
                    AppID: "",
                    RedirectionLink: resData[j].RedirectionLink
                  })
                }
              }
            })
          }

        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArIF", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  } else {
    // alert(`acting for ${Actingfor} is null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];
        if (res.length > 0) {
          for (let i = 0; i < res.length; i++) {
            if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
              await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
                .then((resData) => {
                  if (resData && resData.length > 0) {
                    for (let j = 0; j < resData.length; j++) {
                      // AllApprovalArr.push(resData[j])
                      // AllApprovalArr.push({
                      //   ID: resData[j].ID,
                      //   RequestID: resData[j].ID,
                      //   ApprovalTitle: resData[j].Title,
                      //   Author: resData[j].Requestor_x0020_Name,
                      //   ProcessName: res[i].ProcessName,
                      //   Created: resData[j].Created,
                      //   Status: resData[j].TaskStatus,
                      //   TaskID: resData[j].MasterID,
                      //   AppID: res[i].AppId,
                      //   RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      // })
                      AllApprovalArr.push({
                        ID: resData[j].ID,
                        RequestID: resData[j].Title,
                        ApprovalTitle: "",
                        Author: resData[j].Requestor_x0020_Name,
                        ProcessName: res[i].ProcessName,
                        Created: resData[j].Created,
                        Status: resData[j].TaskStatus,
                        TaskID: resData[j].ID,
                        AppID: res[i].AppId,
                        RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      })
                    }
                  }
                  // Handle the combined data here
                  console.log("Final combined data:", resData);
                })
                .catch((error) => {
                  console.error("Error fetching data:", error);
                });


            } else {
              await getMyApprovalsdata(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
                if (resData && resData.length > 0) {
                  console.log("resDataresDataresDataresData", resData);
                  for (let j = 0; j < resData.length; j++) {
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].RequestID,
                      ApprovalTitle: resData[j].ApprovalTitle,
                      Author: resData[j].Author,
                      ProcessName: resData[j].ProcessName,
                      Created: resData[j].Created,
                      Status: resData[j].Status,
                      TaskID: "",
                      AppID: "",
                      RedirectionLink: resData[j].RedirectionLink
                    })
                  }
                }
              })
            }
          }
        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArrelse", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  }

}

export const getMyRequestsdata = async (_sp, listName) => {
  let arr = []
  let currentUser;
  await _sp.web.currentUser()
    .then(user => {
      console.log("user", user);
      currentUser = user.Email; // Get the current user's Email
    })
    .catch(error => {
      console.error("Error fetching current user: ", error);
      return [];
    });

  if (!currentUser) return arr; // Return empty array if user fetch failed

  await _sp.web.lists.getByTitle(listName).items
    .select("*,Author/ID,Author/Title,Author/EMail").expand("Author")
    .filter(`Author/EMail eq '${currentUser}'`)
    .orderBy("Created", false).getAll()
    .then((res) => {
      console.log(`--MyRequest${listName}`, res);
      arr = res
      // arr = res.filter(item => 
      //     // Include public groups or private groups where the current user is in the InviteMembers array
      //     item.GroupType === "Public" || 
      //     (item.GroupType === "Private" && item.InviteMemebers && item.InviteMemebers.some(member => member.Id === currentUser))
      //   );
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

export const getRequestListsData = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("AllRequestLists").items.orderBy("Created", false).getAll()
    .then((res) => {
      console.log("AllRequestLists", res);
      let AllRequestArr = [];

      for (let i = 0; i < res.length; i++) {
        getMyRequestsdata(_sp, res[i].Title).then((resData) => {
          for (let j = 0; j < resData.length; j++) {
            AllRequestArr.push(resData[j])
          }

        })
      }
      console.log("AllRequestArr", AllRequestArr);
      arr = AllRequestArr;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}