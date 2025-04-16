//import * as React from 'react';
//import styles from './ArgAutomation.module.scss';
import { IDashboardTestProps } from './IDashboardTestProps';

import React, { useEffect, useState } from "react";
import { getSP } from "../../businessApps/loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
//import "../../../CustomCss/mainCustom.scss";
//import "../../discussionForum/components/DiscussionForum.scss";
//import "../../verticalSideBar/components/VerticalSidebar.scss";
//import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../GlobalContext/context";
import Provider from "../../../GlobalContext/provider";
//import { useMediaQuery } from "react-responsive";
import context from "../../../GlobalContext/context";
//import CustomCarousel from "../../../CustomJSComponents/carousel/CustomCarousel";
//import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
//import "./DashboardTest.scss";
let siteID: any;
let response: any;
const HelloWorldContext = ({ props }: any) => {
  const sp: SPFI = getSP();
  const { useHide }: any = React.useContext(UserContext);
  const [RedirectUrl, setRedirectUrl] = useState("");
  const elementRef = React.useRef<HTMLDivElement>(null);
  const { setHide }: any = context;
  React.useEffect(() => {
    ApiCall();

    // const currentUrl = window.location.href;
    // const rootUrl = "https://alrostamanigroupae.sharepoint.com";
    // const redirectUrl = "https://alrostamanigroupae.sharepoint.com/sites/Intranet/SitePages/Home.aspx";

    // if (currentUrl === rootUrl || currentUrl === rootUrl + "/") {
    //   window.location.href = redirectUrl;
    // }

  }, [useHide]);
  // Media query to check if the screen width is less than 768px
  // useEffect(() => {
  //   const currentUrl = window.location.href;
  //   const rootUrl = "https://alrostamanigroupae.sharepoint.com";
  //   const redirectUrl = "https://alrostamanigroupae.sharepoint.com/sites/Intranet/SitePages/Home.aspx";

  //   if (currentUrl === rootUrl || currentUrl === rootUrl + "/") {
  //     window.location.href = redirectUrl;
  //   }
  // }, []);

  const siteUrl = props.siteUrl;
  // useEffect(() => {
  //   ApiCall();
  // }, [])
  const ApiCall = async () => {
    debugger
    let listTitle = 'ConfigurationURLList'
    let arr: any;
    // let CurrentsiteID = props.context.pageContext.site.id;
    // siteID = CurrentsiteID;
    let SiteBaseURL = "https://alrostamanigroupae.sharepoint.com/"
    let apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listTitle}')/items?$select=*&$filter=${`Title eq 'ProductionURL'`}`;
    //response = await sp.web.lists.getByTitle(listTitle).items.select('*')();
    console.log("ressefesre", response, apiUrl);
    try {
      console.log("apiUrlapiUrl", apiUrl);
      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
        },
      });
      console.log("respopopopop", response);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }

      // Wait for the response to be converted to JSON
      const data = await response.json();

      // Immediately handle data
      console.log('List Items ellllllllnew:', data.value);
      let redirectUrl = "";
      // You can now manipulate or return the data as needed
      if (data.value.length > 0) {
        arr = data.value;
        redirectUrl = data.value[0].URL;
      } else {
        arr = []
      }
      const currentUrl = window.location.href;
      const matchUrl = "https://alrostamanigroupae.sharepoint.com";
      if (currentUrl.includes(matchUrl)) {
        window.location.href = redirectUrl;
      }
      //return arr;  // Directly returning the array if needed
    } catch (error) {
      console.error('Error fetching data:', error);
    }
  }
  // const GotoNextPage = (item: any) => {
  //   console.log("item-->>>>", item)
  //   debugger
  //   window.location.href = item.Url;
  //   // const encryptedId = encryptId(String(item.ID));
  //   // sessionStorage.setItem("mediaId", encryptedId);
  //   // sessionStorage.setItem("dataID", item.Id)
  //   // window.location.href = `${siteUrl}/SitePages/Mediadetails.aspx`;
  // };
  // const Breadcrumb = [
  //   {
  //     "MainComponent": "Home",
  //     "MainComponentURl": `${siteUrl}/SitePages/Dashboard.aspx`
  //   },
  //   {
  //     "ChildComponent": "Business App",
  //     "ChildComponentURl": `${siteUrl}/SitePages/DashboardTest.aspx`
  //   }
  // ]
  // const handleRedirect = (link: any) => {
  //   console.log(link, "----link");
  //   window.location.href = link;
  // };
  return (
    <div id="wrapper" ref={elementRef}>


    </div>

  );
};

const DashboardTest: React.FC<IDashboardTestProps> = (props) => (

  <Provider>
    <HelloWorldContext props={props} />

  </Provider>

)

export default DashboardTest;



