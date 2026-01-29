import * as React from 'react';
//import styles from './DashboardReport.module.scss';
import type { IDashboardReportProps } from './IDashboardReportProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import "./DashboardReportcss.scss";
import Provider from '../../../GlobalContext/provider';
export interface IDashboardReportState {
  RowSelected: boolean;
  isChartReady: boolean;
}
declare global {
  interface Window {
    google: any; // or better: google: typeof google.charts;
  }
}
// require("../assets/assets/css/bootstrap.min.css");
//       require("../assets/assets/css/app.min.css");
//       require("../assets/css/icons.min.css");
//       require("../assets/assets/libs/flatpickr/flatpickr.min.css");
//       require("../assets/assets/libs/selectize/css/selectize.bootstrap3.css");
//       require("../assets/assets/js/head.js");
//       require("../assets/assets/libs/apexcharts/apexcharts.min.js");
//       require("../assets/assets/js/pages/dashboard-1.init.js");
export default class DashboardReport extends React.Component<IDashboardReportProps, IDashboardReportState> {
  public constructor(props: IDashboardReportProps | Readonly<IDashboardReportProps>) {
    super(props);
    //Load CSS files
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/css/app.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/css/icons.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/flatpickr/flatpickr.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/selectize/css/selectize.bootstrap3.css");

    // Load JS files in correct order using Promise.all
    Promise.all([
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/js/head.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/flatpickr/flatpickr.min.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/apexcharts/apexcharts.min.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/selectize/js/standalone/selectize.min.js")
    ]).then(() => {
      // All dependencies loaded, now load dashboard init script
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/js/pages/dashboard-1.init.js");
    }).catch(error => {
      console.error("Script loading failed:", error);
    });

    SPComponentLoader.loadScript('https://www.gstatic.com/charts/loader.js', {
      globalExportsName: 'google',
    }).then(() => {
      window['google'].charts.load('current', { packages: ['corechart'] });
      window['google'].charts.setOnLoadCallback(this.drawChart);
      window.google.charts.load('current', {
        packages: ['corechart', 'bar'],
      });
      window.google.charts.setOnLoadCallback(this.drawAnnotations);
    });
    this.state = {
      RowSelected: false,
      isChartReady: false,
      // Initialize any state variables if needed
    };
  }
  async componentDidMount() {
    console.log("test");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/css/app.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/css/icons.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/flatpickr/flatpickr.min.css");
    SPComponentLoader.loadCss("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/selectize/css/selectize.bootstrap3.css");

    // Load JS files in correct order using Promise.all
    Promise.all([
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/js/head.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/flatpickr/flatpickr.min.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/apexcharts/apexcharts.min.js"),
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/libs/selectize/js/standalone/selectize.min.js")
    ]).then(() => {
      // All dependencies loaded, now load dashboard init script
      SPComponentLoader.loadScript("https://officeindia.sharepoint.com/sites/multiverseintranetportal/assets/js/pages/dashboard-1.init.js");
    }).catch(error => {
      console.error("Script loading failed:", error);
    });

    SPComponentLoader.loadScript('https://www.gstatic.com/charts/loader.js', {
      globalExportsName: 'google',
    }).then(() => {
      window['google'].charts.load('current', { packages: ['corechart'] });
      window['google'].charts.setOnLoadCallback(this.drawChart);
      window.google.charts.load('current', {
        packages: ['corechart', 'bar'],
      });
      window.google.charts.setOnLoadCallback(this.drawAnnotations);
    });
  }
  private drawChart = (): void => {
    const data = window['google'].visualization.arrayToDataTable([
      ['Task', 'Hours per Day'],
      ['Approved', 5],
      ['Under Review', 5],
      ['Not Approved', 7],
    ]);

    const options = {
      title: '',
      backgroundColor: 'transparent',
      pieHole: 0.4,
      colors: ['#2c9942', '#52565a', '#ffb748'],
      legend: {
        position: 'bottom',
      },
    };

    const chart = new window['google'].visualization.PieChart(document.getElementById('donutchart'));
    chart.draw(data, options);

    this.setState({ isChartReady: true });
  };

  private drawAnnotations = (): void => {
    const data = new window.google.visualization.DataTable();
    data.addColumn('string', 'Month');
    data.addColumn('number', 'Total Document');
    data.addColumn({ type: 'string', role: 'annotation' });
    data.addColumn('number', 'Approved');
    data.addColumn({ type: 'string', role: 'annotation' });
    data.addColumn('number', 'Under Review');
    data.addColumn({ type: 'string', role: 'annotation' });
    data.addColumn('number', 'Not Approved');
    data.addColumn({ type: 'string', role: 'annotation' });

    data.addRows([
      ['Jan', 25, '25', 48, '48', 60, '60', 80, '80'],
      ['Feb', 35, '35', 63, '63', 80, '80', 17, '17'],
      ['Mar', 20, '20', 35, '35', 50, '50', 30, '30'],
      ['Apr', 25, '25', 48, '48', 60, '60', 80, '80'],
      ['May', 35, '35', 63, '63', 80, '80', 17, '17'],
      ['Jun', 20, '20', 35, '35', 50, '50', 30, '30'],
    ]);

    const options = {
      title: '',
      backgroundColor: 'transparent',
      annotations: {
        alwaysOutside: true,
        textStyle: {
          fontSize: 12,
          color: '#000',
          auraColor: 'none',
        },
      },
      hAxis: {
        title: '',
        gridlines: {
          color: 'transparent',
        },
      },
      vAxis: {
        title: 'Count of Doc Status',
        gridlines: {
          color: 'transparent',
        },
      },
      legend: { position: 'bottom', maxLines: 4 },
      bar: { groupWidth: '75%' },
      colors: ['#118dff', '#2c9942', '#52565a', '#ffb748'],
    };

    const chart = new window.google.visualization.ColumnChart(
      document.getElementById('chart_div')
    );
    chart.draw(data, options);
  };
  private backclick =async () => {
    await this.componentDidMount();
    this.setState({ RowSelected: false });
  }
  private handleSelectChange = async (event: any) => {
    this.setState({ RowSelected: true });
  };
  public render(): React.ReactElement<IDashboardReportProps> {
    let Dashboardcomponent;
    {
      Dashboardcomponent =
        (
          <><div id="wrapper">
            <div style={{ background: '#fafbfc' }} className="content-page">


              <div className="navbar-custom">
                <div className="topbar">
                  <div className="topbar-menu d-flex align-items-center gap-1">

                    {/* <!-- Topbar Brand Logo --> */}
                    <div className="logo-box">
                      {/* <!-- Brand Logo Light --> */}
                      <a onClick={() => this.backclick()} className="logo-light">
                        <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="logo" className="logo-lg" />
                        <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="small logo" className="logo-sm" />
                      </a>

                      {/* <!-- Brand Logo Dark --> */}
                      <a onClick={() => this.backclick()} className="logo-dark">
                        <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="dark logo" className="logo-lg" />
                        <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="small logo" className="logo-sm" />
                      </a>
                    </div>

                    {/* <!-- Sidebar Menu Toggle Button --> */}
                    <button className="button-toggle-menu">
                      <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="dark logo" className="logo-lg" />
                    </button>

                    {/* <!-- Dropdown Menu --> */}
                    <div style={{ marginLeft: '180px' }} className="dropdown d-none d-xl-block">
                      <a className="nav-link dropdown-toggle waves-effect waves-light text-dark font-16 fw-bold" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                        Doesn't have a single universally recognized

                      </a>

                    </div>

                    {/* <!-- Mega Menu Dropdown --> */}

                  </div>

                  <ul className="topbar-menu d-flex align-items-center">
                    {/* <!-- Topbar Search Form --> */}





                    {/* <!-- App Dropdown --> */}
                    <li className="dropdown d-none d-md-inline-block">
                      <a className="nav-link dropdown-toggle waves-effect waves-light arrow-none" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                        <img src={require("../assets/icon/CombinedShape.png")} alt="dark logo" className="logo-lg" />
                      </a>
                      <div className="dropdown-menu dropdown-menu-end dropdown-menu-animated dropdown-lg p-0">

                        <div className="p-2">
                          <div className="row g-0">
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/slack.png")} alt="slack" />
                                <span>Slack</span>
                              </a>
                            </div>
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/github.png")} alt="Github" />
                                <span>GitHub</span>
                              </a>
                            </div>
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/dribbble.png")} alt="dribbble" />
                                <span>Dribbble</span>
                              </a>
                            </div>
                          </div>

                          <div className="row g-0">
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/bitbucket.png")} alt="bitbucket" />
                                <span>Bitbucket</span>
                              </a>
                            </div>
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/dropbox.png")} alt="dropbox" />
                                <span>Dropbox</span>
                              </a>
                            </div>
                            <div className="col">
                              <a className="dropdown-icon-item" href="#">
                                <img src={require("../assets/images/brands/g-suite.png")} alt="G Suite" />
                                <span>G Suite</span>
                              </a>
                            </div>
                          </div>
                        </div>
                      </div>
                    </li>







                    <li className="dropdown">
                      <a className="nav-link dropdown-toggle nav-user me-0 waves-effect waves-light" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                        <img src={require("../assets/images/users/user-3.jpg")} alt="user-image" className="rounded-circle" />
                        <span className="ms-1 d-none d-md-inline-block">
                          Hi Ali Rashid
                        </span>
                      </a>
                      <div className="dropdown-menu dropdown-menu-end profile-dropdown ">

                        <div className="dropdown-header noti-title">
                          <h6 className="text-overflow m-0">Welcome !</h6>
                        </div>


                        <a href="javascript:void(0);" className="dropdown-item notify-item">
                          <i className="fe-user"></i>
                          <span>My Account</span>
                        </a>


                        <a href="javascript:void(0);" className="dropdown-item notify-item">
                          <i className="fe-settings"></i>
                          <span>Settings</span>
                        </a>


                        <a href="javascript:void(0);" className="dropdown-item notify-item">
                          <i className="fe-lock"></i>
                          <span>Lock Screen</span>
                        </a>

                        <div className="dropdown-divider"></div>


                        <a href="javascript:void(0);" className="dropdown-item notify-item">
                          <i className="fe-log-out"></i>
                          <span>Logout</span>
                        </a>

                      </div>
                    </li>

                    <li>
                      <a className="nav-link waves-effect waves-light" data-bs-toggle="offcanvas" href="#theme-settings-offcanvas">
                        <i className="fe-settings font-22"></i>
                      </a>
                    </li>
                  </ul>
                </div>
              </div>
              {/* <!-- ========== Topbar End ========== --> */}

              <div className="content">

                {/* <!-- Start Content--> */}
                <div className="container-fluid mt-4">
                  <div style={{
                    background: '#edf5f1',
                    borderRadius: '10px'
                  }}
                    className="mt-4 mb-2 p-3">
                    <div className="row">
                      <div className="col-sm-3">
                        <div className="mb-2">
                          <label htmlFor="example-select" className="form-label">Year</label>
                          <select className="form-select" id="example-select">
                            <option>All</option>
                            <option>2025</option>
                            <option>2024</option>
                            <option>2023</option>
                            <option>2022</option>
                            <option>2021</option>
                            <option>2020</option>
                            <option>2019</option>
                            <option>2018</option>
                          </select>
                        </div>
                      </div>
                      <div className="col-sm-3">
                        <div className="mb-2">
                          <label htmlFor="example-select" className="form-label">Month</label>
                          <select className="form-select" id="example-select">
                            <option>All</option>
                            <option>Jan</option>
                            <option>Fab</option>
                            <option>Mar</option>
                            <option>Apr</option>
                            <option>May</option>
                            <option>Jun</option>
                            <option>Jul</option>
                            <option>Aug</option>
                            <option>Sep</option>
                            <option>Oct</option>
                            <option>Nov</option>
                            <option>Dec</option>
                          </select>
                        </div>
                      </div>
                      <div className="col-sm-3">
                        <div className="mb-2">
                          <label htmlFor="example-select" className="form-label">Document Status</label>
                          <select className="form-select" id="example-select">
                            <option>All</option>
                            <option>Approved</option>
                            <option>Under Review</option>
                            <option>Not Approved</option>

                          </select>
                        </div>
                      </div>
                      <div className="col-sm-3">
                        <div className="mb-2">
                          <label htmlFor="example-select" className="form-label">Entity</label>
                          <select className="form-select" id="example-select" onChange={this.handleSelectChange}>
                            <option>All</option>
                            <option>Human Resources	</option>
                            <option>Finance</option>
                            <option>IT Department	</option>
                            <option>Marketing</option>
                            <option>Security</option>
                          </select>
                        </div>
                      </div>


                    </div>



                    <div className="card shadow-none mb-0">
                      <div className="card-body p-2">
                        <div className="row">
                          <div className="col-md-6 col-xl-3">
                            <div style={{
                              boxShadow: 'none',
                              borderRight: '1px solid #c8c8c8',
                              paddingRight: '40px'
                            }}
                              className="widget-rounded-circle card mb-0">
                              <div className="card-body p-0">
                                <div className="row">
                                  <div className="col-4">
                                    <div className="">
                                      <img src={require("../assets/icon/f1.png")} />
                                    </div>
                                  </div>
                                  <div className="col-8">
                                    <div className="text-end">
                                      <h3 className="text-dark fw-bold mt-1 mb-0"><span data-plugin="counterup">371</span></h3>
                                      <p className="text-dark mb-1 text-truncate">Total Documents</p>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          <div className="col-md-6 col-xl-3">
                            <div style={{
                              boxShadow: 'none',
                              borderRight: '1px solid #c8c8c8',
                              paddingRight: '40px'
                            }}
                              className="widget-rounded-circle card mb-0">
                              <div className="card-body p-0">
                                <div className="row">
                                  <div className="col-4">
                                    <div className="">
                                      <img src={require("../assets/icon/f2.png")} />
                                    </div>
                                  </div>
                                  <div className="col-8">
                                    <div className="text-end">
                                      <h3 className="text-dark fw-bold mt-1 mb-0"><span data-plugin="counterup">200</span></h3>
                                      <p className="text-dark mb-1 text-truncate">Approved</p>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          <div className="col-md-6 col-xl-3">
                            <div style={{
                              boxShadow: 'none',
                              borderRight: '1px solid #c8c8c8',
                              paddingRight: '40px'
                            }}
                              className="widget-rounded-circle card mb-0">
                              <div className="card-body p-0">
                                <div className="row">
                                  <div className="col-4">
                                    <div className="">
                                      <img src={require("../assets/icon/f3.png")} />
                                    </div>
                                  </div>
                                  <div className="col-8">
                                    <div className="text-end">
                                      <h3 className="text-dark fw-bold mt-1 mb-0"><span data-plugin="counterup">50</span></h3>
                                      <p className="text-dark mb-1 ">Under Review</p>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                          <div className="col-md-6 col-xl-3">
                            <div style={{ boxShadow: 'none', borderRight: '0px solid #c8c8c8', paddingRight: '20px' }}
                              className="widget-rounded-circle card mb-0">
                              <div className="card-body p-0">
                                <div className="row">
                                  <div className="col-4">
                                    <div className="">
                                      <img src={require("../assets/icon/f4.png")} />
                                    </div>
                                  </div>
                                  <div className="col-8">
                                    <div className="text-end">
                                      <h3 className="text-dark fw-bold mt-1 mb-0"><span data-plugin="counterup">71</span></h3>
                                      <p className="text-dark mb-1 ">Not Approved</p>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                        </div>
                      </div>

                    </div>
                  </div>


                  <div className="row">
                    <div className="col-lg-7">
                      <div className="card mb-2">
                        <div className="card-body pb-2">


                          <h4 className="header-title mb-3 fw-bold">Count of Document Status by Month</h4>

                          <div dir="ltr">
                            <div style={{ height: '300px', width: '850px' }} id="chart_div"></div>
                            {/* <img src={require("../assets/icon/barchart.png")} width="100%" /> */}
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="col-lg-5">
                      <div className="row">
                        <div style={{ height: '80px' }} className="col-sm-4 p-0">
                          <div className="first-bg">
                            <h3 className="mb-0 mt-0 fw-bold">31</h3>
                            <p className="mb-0 font-12">Under Review by</p>
                            <p className="mb-0 font-12">7 Days</p>

                          </div>
                        </div>
                        <div style={{ height: '80px' }} className="col-sm-4 p-0">
                          <div className="first-bg2">
                            <h3 className="mb-0 mt-0 fw-bold">48</h3>
                            <p className="mb-0 font-12">Under Review by</p>
                            <p className="mb-0 font-12">8-14 Days</p>

                          </div>
                        </div>

                        <div style={{ height: '80px' }} className="col-sm-4 p-0">
                          <div className="first-bg3">
                            <h3 className="mb-0 mt-0 fw-bold">41</h3>
                            <p className="mb-0 font-12">Under Review by</p>
                            <p className="mb-0 font-12">15 Days</p>

                          </div>
                        </div>
                      </div>

                      <div className="card mt-2 mb-2">
                        <div className="card-body p-1 pb-0">
                          <div className="table-responsive">
                            <table
                              style={{
                                borderCollapse: "separate",
                                borderSpacing: "0 5px",
                                width: "100%",
                              }}
                              className="table table-borderless table-nowrap table-hover table-centered m-0"
                            >
                              <thead>
                                <tr>
                                  <th
                                    className='custom-th2'
                                  >
                                    Document Name
                                  </th>
                                  <th
                                    className='custom-th2'
                                  >
                                    7 Days
                                  </th>
                                  <th
                                    className='custom-th2'
                                  >
                                    8–14 Days
                                  </th>
                                  <th
                                    className='custom-th2'
                                  >
                                    15 Days
                                  </th>
                                </tr>
                              </thead>

                              <tbody className="mt-2">
                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "20px" }} className="text-dark">
                                    <img src={require("../assets/icon/pdf.png")} alt="pdf" /> <span className="text-dark">Nexus Glob.pdf</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">14</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">15</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">5</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "20px" }} className="text-dark">
                                    <img src={require("../assets/icon/doc.png")} alt="doc" /> <span className="text-dark">Nexus Glob.doc</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">2</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">3</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "20px" }} className="text-dark">
                                    <img src={require("../assets/icon/word.png")} alt="word" /> <span className="text-dark">Nexus Glob.word</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">5</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">12</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "20px" }} className="text-dark">
                                    <img src={require("../assets/icon/excel.png")} alt="excel" /> <span className="text-dark">Nexus Glob.Exe</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">5</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10</td>
                                </tr>

                                <tr>
                                  <td style={{ paddingLeft: "20px" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">Total</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">31</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">48</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">41</span>
                                  </td>
                                </tr>
                              </tbody>
                            </table>

                          </div>

                        </div>
                      </div>

                    </div>

                  </div>
                  <div className="row">
                    <div className="col-lg-5">
                      <div className="card mt-2 mb-0">
                        <div className="card-body pb-0">


                          <h4 className="header-title mb-3 fw-bold">Overall Status</h4>

                          <div dir="ltr">
                            <div style={{ height: '300px', width: '800px' }} id="donutchart"></div>
                            {/* <img src={require("../assets/icon/barchart.png")} width="100%" /> */}
                          </div>
                        </div>
                      </div>

                    </div>

                    <div className="col-lg-7">

                      <div className="card mt-2 mb-0">
                        <div className="card-body p-2">
                          <div className="table-responsive">
                            <table style={{ borderCollapse: 'separate', borderSpacing: '0 5px', width: '100%' }}
                              className="table table-border table-nowrap table-hover table-centered m-0">
                              <thead>
                                <tr>
                                  <th
                                  // className='custom-th3'
                                  >
                                    Entity
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Approved
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Not Approved
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Under Review
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Returned to Level 3
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Returned to Level 2
                                  </th>
                                  <th
                                  // style={{
                                  //   background: "#2c9942",
                                  //   color: "#fff",
                                  //   padding: "4px",
                                  //   fontSize: "14px",
                                  //   fontWeight: "400",
                                  //   borderLeft: "1px solid #fff",
                                  //   paddingLeft: "10px",
                                  //   paddingRight: "10px",
                                  //   textAlign: "center",
                                  // }}
                                  >
                                    % Returned to Level 1
                                  </th>
                                </tr>
                              </thead>

                              <tbody className="mt-2">
                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a onClick={(e: any) => this.handleSelectChange(e)} style={{ cursor: 'pointer' }}><span className="text-dark">Human Resources</span></a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">2%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">15%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">5%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">15%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">5%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">3%</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a onClick={(e: any) => this.handleSelectChange(e)}><span className="text-dark" style={{ cursor: 'pointer' }}>Finance</span></a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">4%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">3%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">3%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">1%</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a onClick={(e: any) => this.handleSelectChange(e)}><span className="text-dark" style={{ cursor: 'pointer' }}>IT Department</span></a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">14%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">15%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">15%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a onClick={(e: any) => this.handleSelectChange(e)}><span className="text-dark" style={{ cursor: 'pointer' }}>Marketing</span></a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">2%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">8%</td>
                                </tr>

                                <tr style={{ borderRadius: "5px" }} className="shadow mt-2">
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a onClick={(e: any) => this.handleSelectChange(e)}><span className="text-dark" style={{ cursor: 'pointer' }}>Security</span></a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                  <td style={{ textAlign: "center" }} className="text-dark">10%</td>
                                </tr>

                                <tr>
                                  <td style={{ paddingLeft: "10px" }} className="text-dark">
                                    <a >
                                      <span style={{ color: "#2c9942" }} className="fw-bold">Total</span>
                                    </a>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">32%</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">56%</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">36%</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">56%</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">36%</span>
                                  </td>
                                  <td style={{ textAlign: "center" }} className="text-dark">
                                    <span style={{ color: "#2c9942" }} className="fw-bold">32%</span>
                                  </td>
                                </tr>
                              </tbody>

                            </table>
                          </div>

                        </div>
                      </div>

                    </div>


                  </div>


                </div>

              </div>
              <footer style={{ backgroundColor: '#BEBFC1', color: '#fff', height: '40px' }}
                className="footer">
                <div className="container-fluid">
                  <div className="row">
                    <div className="col-md-12 text-center">
                      2025 © alrostamanigroup
                    </div>

                  </div>
                </div>
              </footer>


            </div>

          </div><div className="offcanvas offcanvas-end right-bar" tabIndex={-1} id="theme-settings-offcanvas">
              <div className="d-flex align-items-center w-100 p-0 offcanvas-header">
                {/* <!-- Nav tabs --> */}
                <ul className="nav nav-tabs nav-bordered nav-justified w-100" role="tablist">
                  <li className="nav-item">
                    <a className="nav-link py-2" data-bs-toggle="tab" href="#chat-tab" role="tab">
                      <i className="mdi mdi-message-text d-block font-22 my-1"></i>
                    </a>
                  </li>
                  <li className="nav-item">
                    <a className="nav-link py-2" data-bs-toggle="tab" href="#tasks-tab" role="tab">
                      <i className="mdi mdi-format-list-checkbox d-block font-22 my-1"></i>
                    </a>
                  </li>
                  <li className="nav-item">
                    <a className="nav-link py-2 active" data-bs-toggle="tab" href="#settings-tab" role="tab">
                      <i className="mdi mdi-cog-outline d-block font-22 my-1"></i>
                    </a>
                  </li>
                </ul>
              </div>

              <div className="offcanvas-body p-3 h-100" data-simplebar>
                {/* <!-- Tab panes --> */}
                <div className="tab-content pt-0">
                  <div className="tab-pane" id="chat-tab" role="tabpanel">

                    <form className="search-bar">
                      <div className="position-relative">
                        <input type="text" className="form-control" placeholder="Search..." />
                        <span className="mdi mdi-magnify"></span>
                      </div>
                    </form>

                    <h6 className="fw-medium mt-2 text-uppercase">Group Chats</h6>

                    <div>
                      <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                        <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-success"></i>
                        <span className="mb-0 mt-1">App Development</span>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                        <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-warning"></i>
                        <span className="mb-0 mt-1">Office Work</span>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                        <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-danger"></i>
                        <span className="mb-0 mt-1">Personal Group</span>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item ps-3 d-block">
                        <i className="mdi mdi-checkbox-blank-circle-outline me-1"></i>
                        <span className="mb-0 mt-1">Freelance</span>
                      </a>
                    </div>

                    <h6 className="fw-medium mt-3 text-uppercase">Favourites <a href="javascript: void(0);" className="font-18 text-danger"><i className="float-end mdi mdi-plus-circle"></i></a></h6>

                    <div>
                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-10.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status online"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Andrew Mackie</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">It will seem like simplified English.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-1.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status away"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Rory Dalyell</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">To an English person, it will seem like simplified</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-9.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status busy"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Jaxon Dunhill</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">To achieve this, it would be necessary.</p>
                            </div>
                          </div>
                        </div>
                      </a>
                    </div>

                    <h6 className="fw-medium mt-3 text-uppercase">Other Chats <a href="javascript: void(0);" className="font-18 text-danger"><i className="float-end mdi mdi-plus-circle"></i></a></h6>

                    <div className="pb-4">
                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-2.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status online"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Jackson Therry</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">Everyone realizes why a new common language.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-4.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status away"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Charles Deakin</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">The languages only differ in their grammar.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-5.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status online"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Ryan Salting</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">If several languages coalesce the grammar of the resulting.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-6.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status online"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Sean Howse</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">It will seem like simplified English.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-7.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status busy"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Dean Coward</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">The new common language will be more simple.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <a href="javascript: void(0);" className="text-reset notification-item">
                        <div className="d-flex align-items-start noti-user-item">
                          <div className="position-relative me-2">
                            <img src={require("../assets/images/users/user-8.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                            <i className="mdi mdi-circle user-status away"></i>
                          </div>
                          <div className="overflow-hidden">
                            <h6 className="mt-0 mb-1 font-14">Hayley East</h6>
                            <div className="font-13 text-muted">
                              <p className="mb-0 text-truncate">One could refuse to pay expensive translators.</p>
                            </div>
                          </div>
                        </div>
                      </a>

                      <div className="text-center mt-3">
                        <a href="javascript:void(0);" className="btn btn-sm btn-white">
                          <i className="mdi mdi-spin mdi-loading me-2"></i>
                          Load more
                        </a>
                      </div>
                    </div>

                  </div>

                  <div className="tab-pane" id="tasks-tab" role="tabpanel">
                    <h6 className="fw-medium p-3 m-0 text-uppercase">Working Tasks</h6>
                    <div className="px-2">
                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          App Development<span className="float-end">75%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-success"
                            role="progressbar"
                            style={{ width: '75%' }}
                            aria-valuenow={75}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>

                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          Database Repair<span className="float-end">37%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-info"
                            role="progressbar"
                            style={{ width: '37%' }}
                            aria-valuenow={37}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>

                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          Backup Create<span className="float-end">52%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-warning"
                            role="progressbar"
                            style={{ width: '52%' }}
                            aria-valuenow={52}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>
                    </div>

                    <h6 className="fw-medium mb-0 mt-4 text-uppercase">Upcoming Tasks</h6>

                    <div>
                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          Sales Reporting<span className="float-end">12%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-danger"
                            role="progressbar"
                            style={{ width: '12%' }}
                            aria-valuenow={12}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>

                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          Redesign Website<span className="float-end">67%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-primary"
                            role="progressbar"
                            style={{ width: '67%' }}
                            aria-valuenow={67}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>

                      <a href="#" className="text-reset item-hovered d-block p-2">
                        <p className="text-muted mb-0">
                          New Admin Design<span className="float-end">84%</span>
                        </p>
                        <div className="progress mt-2" style={{ height: '4px' }}>
                          <div
                            className="progress-bar bg-success"
                            role="progressbar"
                            style={{ width: '84%' }}
                            aria-valuenow={84}
                            aria-valuemin={0}
                            aria-valuemax={100}
                          ></div>
                        </div>
                      </a>
                    </div>

                    <div className="p-3 mt-2 d-grid">
                      <a href="#" className="btn btn-success waves-effect waves-light">
                        Create Task
                      </a>
                    </div>
                  </div>


                  <div className="tab-pane active" id="settings-tab" role="tabpanel">
                    <div className="mt-n3">
                      <h6 className="fw-medium py-2 px-3 font-13 text-uppercase bg-light mx-n3 mt-n3 mb-3">
                        <span className="d-block py-1">Theme Settings</span>
                      </h6>
                    </div>

                    <div className="alert alert-warning" role="alert">
                      <strong>Customize </strong> the overall color scheme, sidebar menu, etc.
                    </div>

                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Color Scheme</h5>

                    <div className="colorscheme-cardradio">
                      <div className="d-flex flex-column gap-2">
                        <div className="form-check form-switch">
                          <input
                            className="form-check-input"
                            type="checkbox"
                            name="data-bs-theme"
                            id="layout-color-light"
                            value="light" />
                          <label className="form-check-label" htmlFor="layout-color-light">
                            Light
                          </label>
                        </div>

                        <div className="form-check form-switch">
                          <input
                            className="form-check-input"
                            type="checkbox"
                            name="data-bs-theme"
                            id="layout-color-dark"
                            value="dark" />
                          <label className="form-check-label" htmlFor="layout-color-dark">
                            Dark
                          </label>
                        </div>
                      </div>
                    </div>

                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Content Width</h5>
                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input
                          className="form-check-input"
                          type="checkbox"
                          name="data-layout-width"
                          id="layout-width-default"
                          value="default" />
                        <label className="form-check-label" htmlFor="layout-width-default">
                          Fluid (Default)
                        </label>
                      </div>

                      <div className="form-check form-switch">
                        <input
                          className="form-check-input"
                          type="checkbox"
                          name="data-layout-width"
                          id="layout-width-boxed"
                          value="boxed" />
                        <label className="form-check-label" htmlFor="layout-width-boxed">
                          Boxed
                        </label>
                      </div>
                    </div>

                    <div id="layout-mode">
                      <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Layout Mode</h5>
                      <div className="d-flex flex-column gap-2">
                        <div className="form-check form-switch">
                          <input
                            className="form-check-input"
                            type="checkbox"
                            name="data-layout-mode"
                            id="layout-mode-default"
                            value="default" />
                          <label className="form-check-label" htmlFor="layout-mode-default">
                            Default
                          </label>
                        </div>

                        <div id="layout-detached">
                          <div className="form-check form-switch">
                            <input
                              className="form-check-input"
                              type="checkbox"
                              name="data-layout-mode"
                              id="layout-mode-detached"
                              value="detached" />
                            <label className="form-check-label" htmlFor="layout-mode-detached">
                              Detached
                            </label>
                          </div>
                        </div>
                      </div>
                    </div>

                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Topbar Color</h5>
                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input
                          className="form-check-input"
                          type="checkbox"
                          name="data-topbar-color"
                          id="topbar-color-light"
                          value="light" />
                        <label className="form-check-label" htmlFor="topbar-color-light">
                          Light
                        </label>
                      </div>

                      <div className="form-check form-switch">
                        <input
                          className="form-check-input"
                          type="checkbox"
                          name="data-topbar-color"
                          id="topbar-color-dark"
                          value="dark" />
                        <label className="form-check-label" htmlFor="topbar-color-dark">
                          Dark
                        </label>
                      </div>

                      <div className="form-check form-switch">
                        <input
                          className="form-check-input"
                          type="checkbox"
                          name="data-topbar-color"
                          id="topbar-color-brand"
                          value="brand" />
                        <label className="form-check-label" htmlFor="topbar-color-brand">
                          Brand
                        </label>
                      </div>
                    </div>

                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Color</h5>
                    <div className="d-flex flex-column gap-2">
                      {["light", "dark", "brand", "gradient"].map((val) => (
                        <div className="form-check form-switch" key={val}>
                          <input
                            className="form-check-input"
                            type="checkbox"
                            name="data-menu-color"
                            id={`leftbar-color-${val}`}
                            value={val} />
                          <label
                            className="form-check-label"
                            htmlFor={`leftbar-color-${val}`}
                          >
                            {val.charAt(0).toUpperCase() + val.slice(1)}
                          </label>
                        </div>
                      ))}
                    </div>

                    <div id="menu-icon-color">
                      <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Icon Color</h5>
                      <div className="d-flex flex-column gap-2">
                        {["light", "dark", "brand", "gradient"].map((val) => (
                          <div className="form-check form-switch" key={val}>
                            <input
                              className="form-check-input"
                              type="checkbox"
                              name="data-two-column-color"
                              id={`twocolumn-menu-color-${val}`}
                              value={val} />
                            <label
                              className="form-check-label"
                              htmlFor={`twocolumn-menu-color-${val}`}
                            >
                              {val.charAt(0).toUpperCase() + val.slice(1)}
                            </label>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div>
                      <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Icon Tone</h5>
                      <div className="d-flex flex-column gap-2">
                        {[
                          { id: "menu-icon-default", label: "Default", value: "default" },
                          { id: "menu-icon-twotone", label: "Twotone", value: "twotones" },
                        ].map((item) => (
                          <div className="form-check form-switch" key={item.id}>
                            <input
                              className="form-check-input"
                              type="checkbox"
                              name="data-menu-icon"
                              id={item.id}
                              value={item.value} />
                            <label className="form-check-label" htmlFor={item.id}>
                              {item.label}
                            </label>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div id="sidebar-size">
                      <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Sidebar Size</h5>
                      <div className="d-flex flex-column gap-2">
                        {[
                          { id: "leftbar-size-default", label: "Default", value: "default" },
                          {
                            id: "leftbar-size-compact",
                            label: "Compact (Medium Width)",
                            value: "compact",
                          },
                          {
                            id: "leftbar-size-small",
                            label: "Condensed (Icon View)",
                            value: "condensed",
                          },
                          {
                            id: "leftbar-size-full",
                            label: "Full Layout",
                            value: "full",
                          },
                          {
                            id: "leftbar-size-fullscreen",
                            label: "Fullscreen Layout",
                            value: "fullscreen",
                          },
                        ].map((item) => (
                          <div className="form-check form-switch" key={item.id}>
                            <input
                              className="form-check-input"
                              type="checkbox"
                              name="data-sidenav-size"
                              id={item.id}
                              value={item.value} />
                            <label className="form-check-label" htmlFor={item.id}>
                              {item.label}
                            </label>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div id="sidebar-user">
                      <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Sidebar User Info</h5>
                      <div className="form-check form-switch">
                        <input
                          type="checkbox"
                          className="form-check-input"
                          name="data-sidebar-user"
                          id="sidebaruser-check" />
                        <label className="form-check-label" htmlFor="sidebaruser-check">
                          Enable
                        </label>
                      </div>
                    </div>
                  </div>

                </div>
              </div>

              <div className="offcanvas-footer border-top py-2 px-2 text-center">
                <div className="d-flex gap-2">
                  <button type="button" className="btn btn-light w-50" id="reset-layout">Reset</button>
                  <a href="https://1.envato.market/uboldadmin" className="btn btn-danger w-50" target="_blank"><i className="mdi mdi-basket me-1"></i> Buy</a>
                </div>
              </div>
            </div></>
        )
    }
    let Selectedcomponent;
    {
      Selectedcomponent = (
        <><div id="wrapper">
          <div style={{ background: '#fafbfc' }} className="content-page">

            {/* <!-- ========== Topbar Start ========== --> */}
            <div className="navbar-custom">
              <div className="topbar">
                <div className="topbar-menu d-flex align-items-center gap-1">

                  {/* <!-- Topbar Brand Logo --> */}
                  <div className="logo-box">
                    {/* <!-- Brand Logo Light --> */}
                    <a onClick={() => this.backclick()} className="logo-light">
                      <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="logo" className="logo-lg" />
                      <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="small logo" className="logo-sm" />
                    </a>

                    {/* <!-- Brand Logo Dark --> */}
                    <a onClick={() => this.backclick()} className="logo-dark">
                      <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="dark logo" className="logo-lg" />
                      <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="small logo" className="logo-sm" />
                    </a>
                  </div>

                  {/* <!-- Sidebar Menu Toggle Button --> */}
                  <button className="button-toggle-menu">
                    <img src={require("../assets/icon/ESSA_Main_Logo.png")} alt="dark logo" className="logo-lg" />
                  </button>

                  {/* <!-- Dropdown Menu --> */}
                  <div style={{ marginLeft: '180px' }} className="dropdown d-none d-xl-block">
                    <a className="nav-link dropdown-toggle waves-effect waves-light text-dark font-16 fw-bold" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                      Doesn't have a single universally recognized

                    </a>

                  </div>

                  {/* <!-- Mega Menu Dropdown --> */}

                </div>

                <ul className="topbar-menu d-flex align-items-center">
                  {/* <!-- Topbar Search Form --> */}





                  {/* <!-- App Dropdown --> */}
                  <li className="dropdown d-none d-md-inline-block">
                    <a className="nav-link dropdown-toggle waves-effect waves-light arrow-none" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                      <img src={require("../assets/icon/CombinedShape.png")} alt="dark logo" className="logo-lg" />
                    </a>
                    <div className="dropdown-menu dropdown-menu-end dropdown-menu-animated dropdown-lg p-0">

                      <div className="p-2">
                        <div className="row g-0">
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/slack.png")} alt="slack" />
                              <span>Slack</span>
                            </a>
                          </div>
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/github.png")} alt="Github" />
                              <span>GitHub</span>
                            </a>
                          </div>
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/dribbble.png")} alt="dribbble" />
                              <span>Dribbble</span>
                            </a>
                          </div>
                        </div>

                        <div className="row g-0">
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/bitbucket.png")} alt="bitbucket" />
                              <span>Bitbucket</span>
                            </a>
                          </div>
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/dropbox.png")} alt="dropbox" />
                              <span>Dropbox</span>
                            </a>
                          </div>
                          <div className="col">
                            <a className="dropdown-icon-item" href="#">
                              <img src={require("../assets/images/brands/g-suite.png")} alt="G Suite" />
                              <span>G Suite</span>
                            </a>
                          </div>
                        </div>
                      </div>
                    </div>
                  </li>






                  {/* <!-- User Dropdown --> */}
                  <li className="dropdown">
                    <a className="nav-link dropdown-toggle nav-user me-0 waves-effect waves-light" data-bs-toggle="dropdown" href="#" role="button" aria-haspopup="false" aria-expanded="false">
                      <img src={require("../assets/images/users/user-3.jpg")} alt="user-image" className="rounded-circle" />
                      <span className="ms-1 d-none d-md-inline-block">
                        Hi Ali Rashid
                      </span>
                    </a>
                    <div className="dropdown-menu dropdown-menu-end profile-dropdown ">

                      <div className="dropdown-header noti-title">
                        <h6 className="text-overflow m-0">Welcome !</h6>
                      </div>


                      <a href="javascript:void(0);" className="dropdown-item notify-item">
                        <i className="fe-user"></i>
                        <span>My Account</span>
                      </a>


                      <a href="javascript:void(0);" className="dropdown-item notify-item">
                        <i className="fe-settings"></i>
                        <span>Settings</span>
                      </a>


                      <a href="javascript:void(0);" className="dropdown-item notify-item">
                        <i className="fe-lock"></i>
                        <span>Lock Screen</span>
                      </a>

                      <div className="dropdown-divider"></div>


                      <a href="javascript:void(0);" className="dropdown-item notify-item">
                        <i className="fe-log-out"></i>
                        <span>Logout</span>
                      </a>

                    </div>
                  </li>

                  {/* <!-- Right Bar offcanvas button (Theme Customization Panel) --> */}
                  <li>
                    <a className="nav-link waves-effect waves-light" data-bs-toggle="offcanvas" href="#theme-settings-offcanvas">
                      <i className="fe-settings font-22"></i>
                    </a>
                  </li>
                </ul>
              </div>
            </div>
            {/* <!-- ========== Topbar End ========== --> */}

            <div className="content">

              {/* <!-- Start Content--> */}
              <div className="container-fluid mt-2">

                {/* <!-- end row--> */}
                <div className="row">
                  <div className="col-sm-6">
                    <div className="page-title-box">
                      <div className="page-title-left">
                        <h4 style={{ lineHeight: '20px' }} className="page-title mb-0 font-18 fw-bold">Total Documents</h4>
                        <ol className="breadcrumb mt-0 mb-0">
                          <li className="breadcrumb-item"><a href="javascript: void(0);" className="text-dark">Dashboard</a></li>

                          <li className="breadcrumb-item text-dark active">Total Documents</li>
                        </ol>
                      </div>

                    </div>
                  </div>

                  <div className="col-sm-6">

                    <div style={{ textAlign: 'right' }} className="pull-right mt-1">
                      <a ><div className="btn btn-secondary" onClick={() => this.backclick()}>Back</div> </a>
                    </div>


                  </div>
                </div>

                <div className="row">


                  <div className="col-lg-12">

                    <div className="card mt-2 mb-0">
                      <div className="card-body p-2">
                        <div className="table-responsive">
                          <table style={{
                            borderCollapse: 'separate',
                            borderSpacing: '0 5px',
                            width: '100%'
                          }}
                            className="table table-border table-nowrap table-hover table-centered m-0">

                            <thead>
                              <tr>
                                <th
                                //className="custom-th"
                                >
                                  Entity
                                </th>

                                <th
                                // style={{
                                //   background: '#2c9942',
                                //   color: '#fff',
                                //   padding: '4px',
                                //   fontSize: '14px',
                                //   fontWeight: 400,
                                //   borderLeft: '1px solid #fff',
                                //   paddingLeft: '10px',
                                //   paddingRight: '10px',
                                //   textAlign: 'center',
                                // }}
                                >
                                  Document Name
                                </th>

                                <th
                                // style={{
                                //   background: '#2c9942',
                                //   color: '#fff',
                                //   padding: '4px',
                                //   fontSize: '14px',
                                //   fontWeight: 400,
                                //   borderLeft: '1px solid #fff',
                                //   paddingLeft: '10px',
                                //   paddingRight: '10px',
                                //   textAlign: 'center',
                                // }}
                                >
                                  Submitted Date
                                </th>

                                <th
                                // style={{
                                //   background: '#2c9942',
                                //   color: '#fff',
                                //   padding: '4px',
                                //   fontSize: '14px',
                                //   fontWeight: 400,
                                //   borderLeft: '1px solid #fff',
                                //   paddingLeft: '10px',
                                //   paddingRight: '10px',
                                //   textAlign: 'center',
                                // }}
                                >
                                  Status
                                </th>

                                <th
                                // style={{
                                //   background: '#2c9942',
                                //   color: '#fff',
                                //   padding: '4px',
                                //   fontSize: '14px',
                                //   fontWeight: 400,
                                //   borderLeft: '1px solid #fff',
                                //   paddingLeft: '10px',
                                //   paddingRight: '10px',
                                //   textAlign: 'center',
                                // }}
                                >
                                  Owner
                                </th>
                              </tr>
                            </thead>

                            <tbody className="mt-2">
                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">Human Resources</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/pdf.png")} /> <span className="text-dark">DD.PFD</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="approved">Approved</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener1.png")} /> <span className="text-dark">Varun Sharma</span>
                                </td>
                              </tr>

                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">Finance</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/word.png")} /> <span className="text-dark">GA.DOC</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="notapproved">Not Approved</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener1.png")} /> <span className="text-dark">Naman Gupta</span>
                                </td>
                              </tr>

                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">IT Department</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/pdf.png")} /> <span className="text-dark">PID.pdf</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="underreview">Under Review</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener3.png")} /> <span className="text-dark">Karan Varma</span>
                                </td>
                              </tr>

                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">Marketing</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/word.png")} /> <span className="text-dark">MAR.DOC</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="approved">Approved</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener4.png")} /> <span className="text-dark">Akash Gupta</span>
                                </td>
                              </tr>

                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">Security</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/pdf.png")} /> <span className="text-dark">SOP.PDF</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="notapproved">Not Approved</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener5.png")} /> <span className="text-dark">Tarun Kumar</span>
                                </td>
                              </tr>

                              <tr style={{ borderRadius: '5px' }} className="shadow mt-2">
                                <td style={{ paddingLeft: '10px' }} className="text-dark">
                                  <span className="text-dark">Operations</span>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/word.png")} /> <span className="text-dark">MOM.DOC</span>
                                </td>
                                <td style={{ textAlign: 'center' }} className="text-dark">04-07-2025</td>
                                <td style={{ textAlign: 'center' }} className="text-dark">
                                  <div className="underreview">Under Review</div>
                                </td>
                                <td style={{ textAlign: 'left', paddingLeft: '40px' }} className="text-dark">
                                  <img src={require("../assets/icon/owener6.png")} /> <span className="text-dark">Vishal Kumar</span>
                                </td>
                              </tr>
                            </tbody>

                          </table>
                        </div>

                      </div>
                    </div>

                  </div>
                  {/* <!-- end col--> */}


                </div>
                {/* <!-- end row -->


<!-- end row --> */}

              </div>

            </div>
            {/* <!-- content -->

<!-- Footer Start --> */}
            <footer style={{ backgroundColor: '#BEBFC1', color: '#fff', height: '40px' }}
              className="footer">
              <div className="container-fluid">
                <div className="row">
                  <div className="col-md-12 text-center">
                    2025 © alrostamanigroup
                  </div>

                </div>
              </div>
            </footer>

            {/* <!-- end Footer --> */}

          </div>

        </div><div className="offcanvas offcanvas-end right-bar" tabIndex={-1} id="theme-settings-offcanvas">
            <div className="d-flex align-items-center w-100 p-0 offcanvas-header">
              {/* <!-- Nav tabs --> */}
              <ul className="nav nav-tabs nav-bordered nav-justified w-100" role="tablist">
                <li className="nav-item">
                  <a className="nav-link py-2" data-bs-toggle="tab" href="#chat-tab" role="tab">
                    <i className="mdi mdi-message-text d-block font-22 my-1"></i>
                  </a>
                </li>
                <li className="nav-item">
                  <a className="nav-link py-2" data-bs-toggle="tab" href="#tasks-tab" role="tab">
                    <i className="mdi mdi-format-list-checkbox d-block font-22 my-1"></i>
                  </a>
                </li>
                <li className="nav-item">
                  <a className="nav-link py-2 active" data-bs-toggle="tab" href="#settings-tab" role="tab">
                    <i className="mdi mdi-cog-outline d-block font-22 my-1"></i>
                  </a>
                </li>
              </ul>
            </div>

            <div className="offcanvas-body p-3 h-100" data-simplebar>
              {/* <!-- Tab panes --> */}
              <div className="tab-content pt-0">
                <div className="tab-pane" id="chat-tab" role="tabpanel">

                  <form className="search-bar">
                    <div className="position-relative">
                      <input type="text" className="form-control" placeholder="Search..." />
                      <span className="mdi mdi-magnify"></span>
                    </div>
                  </form>

                  <h6 className="fw-medium mt-2 text-uppercase">Group Chats</h6>

                  <div>
                    <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                      <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-success"></i>
                      <span className="mb-0 mt-1">App Development</span>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                      <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-warning"></i>
                      <span className="mb-0 mt-1">Office Work</span>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item ps-3 mb-2 d-block">
                      <i className="mdi mdi-checkbox-blank-circle-outline me-1 text-danger"></i>
                      <span className="mb-0 mt-1">Personal Group</span>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item ps-3 d-block">
                      <i className="mdi mdi-checkbox-blank-circle-outline me-1"></i>
                      <span className="mb-0 mt-1">Freelance</span>
                    </a>
                  </div>

                  <h6 className="fw-medium mt-3 text-uppercase">Favourites <a href="javascript: void(0);" className="font-18 text-danger"><i className="float-end mdi mdi-plus-circle"></i></a></h6>

                  <div>
                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-10.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status online"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Andrew Mackie</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">It will seem like simplified English.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-1.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status away"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Rory Dalyell</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">To an English person, it will seem like simplified</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-9.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status busy"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Jaxon Dunhill</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">To achieve this, it would be necessary.</p>
                          </div>
                        </div>
                      </div>
                    </a>
                  </div>

                  <h6 className="fw-medium mt-3 text-uppercase">Other Chats <a href="javascript: void(0);" className="font-18 text-danger"><i className="float-end mdi mdi-plus-circle"></i></a></h6>

                  <div className="pb-4">
                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-2.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status online"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Jackson Therry</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">Everyone realizes why a new common language.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-4.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status away"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Charles Deakin</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">The languages only differ in their grammar.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-5.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status online"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Ryan Salting</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">If several languages coalesce the grammar of the resulting.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-6.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status online"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Sean Howse</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">It will seem like simplified English.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-7.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status busy"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Dean Coward</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">The new common language will be more simple.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset notification-item">
                      <div className="d-flex align-items-start noti-user-item">
                        <div className="position-relative me-2">
                          <img src={require("../assets/images/users/user-8.jpg")} className="rounded-circle avatar-sm" alt="user-pic" />
                          <i className="mdi mdi-circle user-status away"></i>
                        </div>
                        <div className="overflow-hidden">
                          <h6 className="mt-0 mb-1 font-14">Hayley East</h6>
                          <div className="font-13 text-muted">
                            <p className="mb-0 text-truncate">One could refuse to pay expensive translators.</p>
                          </div>
                        </div>
                      </div>
                    </a>

                    <div className="text-center mt-3">
                      <a href="javascript:void(0);" className="btn btn-sm btn-white">
                        <i className="mdi mdi-spin mdi-loading me-2"></i>
                        Load more
                      </a>
                    </div>
                  </div>

                </div>

                <div className="tab-pane" id="tasks-tab" role="tabpanel">
                  <h6 className="fw-medium p-3 m-0 text-uppercase">Working Tasks</h6>
                  <div className="px-2">
                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">App Development<span className="float-end">75%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div
                          className="progress-bar bg-success"
                          role="progressbar"
                          style={{ width: '75%' }}
                          aria-valuenow={75}
                          aria-valuemin={0}
                          aria-valuemax={100}
                        ></div>

                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">Database Repair<span className="float-end">37%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div className="progress-bar bg-info" role="progressbar" style={{ width: '37%' }} aria-valuenow={37} aria-valuemin={0} aria-valuemax={100}></div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">Backup Create<span className="float-end">52%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div className="progress-bar bg-warning" role="progressbar" style={{ width: '52%' }} aria-valuenow={52} aria-valuemin={0} aria-valuemax={100}></div>
                      </div>
                    </a>
                  </div>

                  <h6 className="fw-medium mb-0 mt-4 text-uppercase">Upcoming Tasks</h6>

                  <div>
                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">Sales Reporting<span className="float-end">12%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div className="progress-bar bg-danger" role="progressbar" style={{ width: '12%' }} aria-valuenow={12} aria-valuemin={0} aria-valuemax={100}></div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">Redesign Website<span className="float-end">67%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div className="progress-bar bg-primary" role="progressbar" style={{ width: '67%' }} aria-valuenow={67} aria-valuemin={0} aria-valuemax={100}></div>
                      </div>
                    </a>

                    <a href="javascript: void(0);" className="text-reset item-hovered d-block p-2">
                      <p className="text-muted mb-0">New Admin Design<span className="float-end">84%</span></p>
                      <div className="progress mt-2" style={{ height: '4px' }}>
                        <div className="progress-bar bg-success" role="progressbar" style={{ width: '84%' }} aria-valuenow={84} aria-valuemin={0} aria-valuemax={100}></div>
                      </div>
                    </a>
                  </div>

                  <div className="p-3 mt-2 d-grid">
                    <a href="javascript: void(0);" className="btn btn-success waves-effect waves-light">Create Task</a>
                  </div>

                </div>

                <div className="tab-pane active" id="settings-tab" role="tabpanel">

                  <div className="mt-n3">
                    <h6 className="fw-medium py-2 px-3 font-13 text-uppercase bg-light mx-n3 mt-n3 mb-3">
                      <span className="d-block py-1">Theme Settings</span>
                    </h6>
                  </div>

                  <div className="alert alert-warning" role="alert">
                    <strong>Customize </strong> the overall color scheme, sidebar menu, etc.
                  </div>

                  <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Color Scheme</h5>

                  <div className="colorscheme-cardradio">
                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-bs-theme" id="layout-color-light" value="light" />
                        <label className="form-check-label" htmlFor="layout-color-light">Light</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-bs-theme" id="layout-color-dark" value="dark" />
                        <label className="form-check-label" htmlFor="layout-color-dark">Dark</label>
                      </div>
                    </div>
                  </div>

                  <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Content Width</h5>
                  <div className="d-flex flex-column gap-2">
                    <div className="form-check form-switch">
                      <input className="form-check-input" type="checkbox" name="data-layout-width" id="layout-width-default" value="default" />
                      <label className="form-check-label" htmlFor="layout-width-default">Fluid (Default)</label>
                    </div>

                    <div className="form-check form-switch">
                      <input className="form-check-input" type="checkbox" name="data-layout-width" id="layout-width-boxed" value="boxed" />
                      <label className="form-check-label" htmlFor="layout-width-boxed">Boxed</label>
                    </div>
                  </div>

                  <div id="layout-mode">
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Layout Mode</h5>

                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-layout-mode" id="layout-mode-default" value="default" />
                        <label className="form-check-label" htmlFor="layout-mode-default">Default</label>
                      </div>


                      <div id="layout-detached">
                        <div className="form-check form-switch">
                          <input className="form-check-input" type="checkbox" name="data-layout-mode" id="layout-mode-detached" value="detached" />
                          <label className="form-check-label" htmlFor="layout-mode-detached">Detached</label>
                        </div>
                      </div>
                    </div>
                  </div>

                  <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Topbar Color</h5>

                  <div className="d-flex flex-column gap-2">
                    <div className="form-check form-switch">
                      <input className="form-check-input" type="checkbox" name="data-topbar-color" id="topbar-color-light" value="light" />
                      <label className="form-check-label" htmlFor="topbar-color-light">Light</label>
                    </div>

                    <div className="form-check form-switch">
                      <input className="form-check-input" type="checkbox" name="data-topbar-color" id="topbar-color-dark" value="dark" />
                      <label className="form-check-label" htmlFor="topbar-color-dark">Dark</label>
                    </div>

                    <div className="form-check form-switch">
                      <input className="form-check-input" type="checkbox" name="data-topbar-color" id="topbar-color-brand" value="brand" />
                      <label className="form-check-label" htmlFor="topbar-color-brand">Brand</label>
                    </div>
                  </div>

                  <div>
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Color</h5>

                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-color" id="leftbar-color-light" value="light" />
                        <label className="form-check-label" htmlFor="leftbar-color-light">Light</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-color" id="leftbar-color-dark" value="dark" />
                        <label className="form-check-label" htmlFor="leftbar-color-dark">Dark</label>
                      </div>
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-color" id="leftbar-color-brand" value="brand" />
                        <label className="form-check-label" htmlFor="leftbar-color-brand">Brand</label>
                      </div>
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-color" id="leftbar-color-gradient" value="gradient" />
                        <label className="form-check-label" htmlFor="leftbar-color-gradient">Gradient</label>
                      </div>
                    </div>
                  </div>

                  <div id="menu-icon-color">
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Icon Color</h5>

                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-two-column-color" id="twocolumn-menu-color-light" value="light" />
                        <label className="form-check-label" htmlFor="twocolumn-menu-color-light">Light</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-two-column-color" id="twocolumn-menu-color-dark" value="dark" />
                        <label className="form-check-label" htmlFor="twocolumn-menu-color-dark">Dark</label>
                      </div>
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-two-column-color" id="twocolumn-menu-color-brand" value="brand" />
                        <label className="form-check-label" htmlFor="twocolumn-menu-color-brand">Brand</label>
                      </div>
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-two-column-color" id="twocolumn-menu-color-gradient" value="gradient" />
                        <label className="form-check-label" htmlFor="twocolumn-menu-color-gradient">Gradient</label>
                      </div>
                    </div>
                  </div>

                  <div>
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Menu Icon Tone</h5>

                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-icon" id="menu-icon-default" value="default" />
                        <label className="form-check-label" htmlFor="menu-icon-default">Default</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-menu-icon" id="menu-icon-twotone" value="twotones" />
                        <label className="form-check-label" htmlFor="menu-icon-twotone">Twotone</label>
                      </div>
                    </div>
                  </div>

                  <div id="sidebar-size">
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Sidebar Size</h5>

                    <div className="d-flex flex-column gap-2">
                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-sidenav-size" id="leftbar-size-default" value="default" />
                        <label className="form-check-label" htmlFor="leftbar-size-default">Default</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-sidenav-size" id="leftbar-size-compact" value="compact" />
                        <label className="form-check-label" htmlFor="leftbar-size-compact">Compact (Medium Width)</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-sidenav-size" id="leftbar-size-small" value="condensed" />
                        <label className="form-check-label" htmlFor="leftbar-size-small">Condensed (Icon View)</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-sidenav-size" id="leftbar-size-full" value="full" />
                        <label className="form-check-label" htmlFor="leftbar-size-full">Full Layout</label>
                      </div>

                      <div className="form-check form-switch">
                        <input className="form-check-input" type="checkbox" name="data-sidenav-size" id="leftbar-size-fullscreen" value="fullscreen" />
                        <label className="form-check-label" htmlFor="leftbar-size-fullscreen">Fullscreen Layout</label>
                      </div>
                    </div>
                  </div>

                  <div id="sidebar-user">
                    <h5 className="fw-medium font-14 mt-4 mb-2 pb-1">Sidebar User Info</h5>

                    <div className="form-check form-switch">
                      <input type="checkbox" className="form-check-input" name="data-sidebar-user" id="sidebaruser-check" />
                      <label className="form-check-label" htmlFor="sidebaruser-check">Enable</label>
                    </div>
                  </div>

                </div>
              </div>
            </div>

            <div className="offcanvas-footer border-top py-2 px-2 text-center">
              <div className="d-flex gap-2">
                <button type="button" className="btn btn-light w-50" id="reset-layout">Reset</button>
                <a href="https://1.envato.market/uboldadmin" className="btn btn-danger w-50" target="_blank"><i className="mdi mdi-basket me-1"></i> Buy</a>
              </div>
            </div>
          </div></>
      )
    }
    return (
      <section>
        {this.state.RowSelected ? Selectedcomponent : Dashboardcomponent}
        {/* <!-- Begin page --> */}


      </section>
    );
  }
}
