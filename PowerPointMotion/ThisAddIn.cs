using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Threading;
using System.Windows.Forms;
using System.Timers;
using Microsoft.Kinect;
using Kinect.Toolbox;

namespace PowerPointMotion
{
    public partial class MotionAddIn
    {
        PowerPoint.SlideShowWindow activeslideShowWindow = null;

        KinectSensor kinectSensor;

        SwipeGestureDetector swipeGestureRecognizerLeftHand;
        SwipeGestureDetector swipeGestureRecognizerRightHand;

        private Skeleton[] skeletons;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideShowBegin += new PowerPoint.EApplication_SlideShowBeginEventHandler(OnPresentationBegin);
            this.Application.SlideShowEnd += new PowerPoint.EApplication_SlideShowEndEventHandler(OnPresentationEnd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        void OnPresentationBegin(PowerPoint.SlideShowWindow Wn)
        {
            activeslideShowWindow = Wn;

            try
            {
                //listen to any status change for Kinects
                KinectSensor.KinectSensors.StatusChanged += Kinects_StatusChanged;

                //loop through all the Kinects attached to this PC, and start the first that is connected without an error.
                foreach (KinectSensor kinect in KinectSensor.KinectSensors)
                {
                    if (kinect.Status == KinectStatus.Connected)
                    {
                        kinectSensor = kinect;
                        break;
                    }
                }

                if (KinectSensor.KinectSensors.Count == 0)
                    MessageBox.Show("No Kinect found");
                else
                    init();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void init()
        {
            if (kinectSensor == null)
                return;

            kinectSensor.SkeletonStream.Enable(new TransformSmoothParameters
            {
                Smoothing = 0.5f,
                Correction = 0.5f,
                Prediction = 0.5f,
                JitterRadius = 0.05f,
                MaxDeviationRadius = 0.04f
            });

            kinectSensor.SkeletonFrameReady += kinectRuntime_SkeletonFrameReady;

            swipeGestureRecognizerLeftHand = new SwipeGestureDetector(/*TODO parameter windowsize*/);
            swipeGestureRecognizerRightHand = new SwipeGestureDetector(/*TODO parameter windowsize*/);
            swipeGestureRecognizerLeftHand.OnGestureDetected += OnGestureDetectedFromLeftHand;
            swipeGestureRecognizerRightHand.OnGestureDetected += OnGestureDetectedFromRightHand;

            kinectSensor.Start();
        }

        void Kinects_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            switch (e.Status)
            {
                case KinectStatus.Connected:
                    if (kinectSensor == null)
                    {
                        kinectSensor = e.Sensor;
                        Initialize();
                    }
                    break;
                case KinectStatus.Disconnected:
                    if (kinectSensor == e.Sensor)
                    {
                        Clean();
                        MessageBox.Show("Kinect was disconnected");
                    }
                    break;
                case KinectStatus.NotReady:
                    break;
                case KinectStatus.NotPowered:
                    if (kinectSensor == e.Sensor)
                    {
                        Clean();
                        MessageBox.Show("Kinect is no more powered");
                    }
                    break;
                default:
                    MessageBox.Show("Unhandled Status: " + e.Status);
                    break;
            }
        }

        private void Clean()
        {
            if (swipeGestureRecognizerLeftHand != null)
            {
                swipeGestureRecognizerLeftHand.OnGestureDetected -= OnGestureDetectedFromLeftHand;
            }

            if (swipeGestureRecognizerRightHand != null)
            {
                swipeGestureRecognizerRightHand.OnGestureDetected -= OnGestureDetectedFromRightHand;
            }
            
            if (kinectSensor != null)
            {
                kinectSensor.SkeletonFrameReady -= kinectRuntime_SkeletonFrameReady;
                kinectSensor.Stop();
                kinectSensor = null;
            }
        }

        void kinectRuntime_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (SkeletonFrame frame = e.OpenSkeletonFrame())
            {
                if (frame == null)
                    return;
                
                frame.GetSkeletons(ref skeletons);

                if (skeletons.All(s => s.TrackingState == SkeletonTrackingState.NotTracked))
                    return;

                foreach (var skeleton in skeletons)
                {
                    if (skeleton.TrackingState != SkeletonTrackingState.Tracked)
                        continue;

                    foreach (Joint joint in skeleton.Joints)
                    {
                        if (joint.TrackingState != JointTrackingState.Tracked)
                            continue;

                        if (joint.JointType == JointType.HandRight)
                        {
                            swipeGestureRecognizerRightHand.Add(joint.Position, kinectSensor);
                        }
                        else if (joint.JointType == JointType.HandLeft)
                        {
                            swipeGestureRecognizerLeftHand.Add(joint.Position, kinectSensor);
                        }
                    }
                }
            }
        }

        void OnPresentationEnd(PowerPoint.Presentation Pr)
        {
            Clean();
            activeslideShowWindow = null;
        }

        void OnGestureDetectedFromRightHand(string gesture)
        {
            switch (gesture)
            {
                case "SwipeToLeft":
                    NextSlide();
                    break;
                case "SwipeToRight":
                    //PrevSlide();
                    break;
                default:
                    break;
            }
        }

        void OnGestureDetectedFromLeftHand(string gesture)
        {
            switch (gesture)
            {
                case "SwipeToLeft":
                    //NextSlide();
                    break;
                case "SwipeToRight":
                    PrevSlide();
                    break;
                default:
                    break;
            }
        }

        void NextSlide()
        {
            if (activeslideShowWindow != null)
            {   
                int slideIndex = activeslideShowWindow.View.Slide.SlideIndex +1;
   
                if (slideIndex <= activeslideShowWindow.Presentation.Slides.Count)
                {
                    try
                    {
                        activeslideShowWindow.View.Next();
                    }
                    catch (Exception e)
                    {
                    }
                } 

            }
        }

        void PrevSlide()
        {
            if (activeslideShowWindow != null)
            {
                int slideIndex = activeslideShowWindow.View.Slide.SlideIndex + 1;

                if (slideIndex <= activeslideShowWindow.Presentation.Slides.Count)
                {
                    try
                    {
                        activeslideShowWindow.View.Previous();
                    }
                    catch (Exception e)
                    {
                    }
                }

            }
        }
    }
}
