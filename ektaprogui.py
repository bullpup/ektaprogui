#!/usr/bin/env python

"""

   Copyright 2010 Julian Hoch

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
   
"""


from Tkconstants import SINGLE, END, DISABLED, HORIZONTAL, BOTTOM, W, X, LEFT, \
    BOTH, RIGHT, N, TOP, NORMAL
from Tkinter import Tk, Frame, Listbox, Button, Label, Entry, IntVar, \
    Checkbutton, Scale, Menu
from thread import allocate_lock
import logging
import serial
import time
import tkMessageBox
import tkSimpleDialog





class EktaproController:
    """ Manages the slide projector devices.  """

    def __init__(self):
        self.activeDevice = None
        self.maxBrightness = 100        
        self.standby = True
        self.devices = []
        self.maxTray= 80
        self.activeIndex = 0
        


    def initDevices(self):
        self.devices = []
        self.maxTray = 0
        for i in range(16):
            try:
                s = serial.Serial(i, timeout=5)
                logging.info("Device on port COM" + str(i + 1) + " found")
                s.write(EktaproCommand(0).statusSystemReturn().toData())
                deviceInfo = s.read(5)
                ed = EktaproDevice(deviceInfo, s, i)
                logger.info(ed)
                logger.debug(ed.getDetails())
                self.devices.append(ed)
                if ed.traySize > self.maxTray:
                    self.maxTray = ed.traySize                    
                
            except serial.SerialException:
                pass
            except IOError:
                logging.error("not a kodakpro device")

        if len(self.devices) > 0:
            self.activeDevice = self.devices[0]
            self.activeIndex = 0
            

    def setActiveDevice(self, items):
        if len(items) == 0:
            self.activeDevice = None            
            return False
        elif len(self.devices) > items[0]:
            self.activeDevice = self.devices[items[0]]
            self.activeIndex = items[0]
            return True

    def resetDevices(self):
        for d in self.devices:           
            d.setStandby(False)
            d.gotoSlide(1)            
            d.setBrightness(0)
        self.standby = False 


    def cleanUp(self):
        for d in self.devices:
            d.resetSystem()
            d.serialDevice.close()


    def getNextDevice(self):
        nextDeviceIndex = self.activeIndex + 1
        if nextDeviceIndex == len(self.devices):
            nextDeviceIndex = 0
        nextDevice = self.devices[nextDeviceIndex]
        return nextDevice


    def getPrevDevice(self):
        prevDeviceIndex = self.activeIndex - 1
        if prevDeviceIndex < 0:
            prevDeviceIndex = len(self.devices) - 1
        prevDevice = self.devices[prevDeviceIndex]
        return prevDevice


    def activateNextDevice(self):
        logger.debug("activating next device")
        self.activeIndex = self.activeIndex + 1
        if self.activeIndex == len(self.devices):
            self.activeIndex = 0
        self.activeDevice = self.devices[self.activeIndex]


    def activatePrevDevice(self):
        logger.debug("activating previous device")
        self.activeIndex = self.activeIndex - 1
        if self.activeIndex < 0:
            self.activeIndex = len(self.devices) - 1
        self.activeDevice = self.devices[self.activeIndex]        
    

    def syncDevices(self):
        for d in self.devices:
            d.sync()

    def toggleStandby(self):       
        self.standby = not self.standby        
        
        for d in self.devices:
            d.setStandby(self.standby)    


class TimerController:
    """ 
    Contains the logic to control the timer and
    fading mechanism.
    """
    

    def __init__(self, controller, gui):
        self.controller = controller
        self.gui = gui
        self.cycle = False
        self.states = {
            0: "IDLE",
            1: "SINGLE_FADING_DOWN",
            2: "SINGLE_FADING_UP",
            3: "DUAL_FADE"            
            }
        
        self.state = 0      
        self.timerCounter = 0
        self.slideshowActive = False
        self.timerActive = False
        self.fadePaused = False
        self.slideshowPaused = False
        self.lock = allocate_lock()
        self.followingDevice = None
        self.slideshowDelay = 5
        self.fadeDelay = 2

    #
    # Public API
    #

    def nextSlide(self):
        activeDevice = self.controller.activeDevice
        nextDevice = self.controller.getNextDevice()
        
        if activeDevice == None:
            return
       
        self.fadeDelay = int(self.gui.fadeInput.get())
        doFade = False if self.fadeDelay == 0 else True

        # Only 1 Projector
        if self.cycle == False or self.isSingleProjector():
            if doFade:                                        
                self.state = 1
                self.goFollowingSlide = activeDevice.gotoNextSlide
                self.lock.acquire()
                if not self.timerActive:
                    self.timerActive = True
                    self.gui.after(50, self.timerEvent)
                self.lock.release()
                return
            else:
                        
                activeDevice.gotoNextSlide()
                if self.slideshowActive:
                    self.lock.acquire()
                    if not self.timerActive:
                        self.timerActive = True
                        self.gui.after(1000 * self.slideshowDelay, \
                                       self.timerEvent)
                    self.lock.release()
                return



        # More than 1 Projector, cycling
        if doFade:            
            self.state = 3
            self.followingDevice = nextDevice
            self.goFollowingSlide = activeDevice.gotoNextSlide
            self.activateFollowingDevice = self.controller.activateNextDevice
            self.lock.acquire()
            if not self.timerActive:
                self.timerActive = True
                self.gui.after(50, self.timerEvent)
            self.lock.release()
            return
        else:  
            activeDevice.setBrightness(0)
            activeDevice.gotoNextSlide()
            nextDevice.setBrightness(self.controller.maxBrightness)
            self.controller.activateNextDevice()
            self.lock.acquire()
            if not self.timerActive:
                self.timerActive = True
                self.gui.after(1000 * self.slideshowDelay, self.timerEvent)
            self.lock.release()
            return
                    

    def previousSlide(self):
        activeDevice = self.controller.activeDevice
        prevDevice = self.controller.getPrevDevice()
        
        if activeDevice == None:
            return
       
        self.fadeDelay = int(self.gui.fadeInput.get())
        doFade = False if self.fadeDelay == 0 else True

        # Only 1 Projector
        if self.cycle == False or self.isSingleProjector():
            if doFade:
                self.state = 1
                self.goFollowingSlide = activeDevice.gotoPrevSlide
                self.lock.acquire()
                if not self.timerActive:
                    self.timerActive = True
                    self.gui.after(50, self.timerEvent)
                self.lock.release()
                return

            else:             
                activeDevice.gotoPrevSlide()
                return

        # More than 1 Projector, cycling
        if doFade:
            self.state = 3
            self.followingDevice = prevDevice
            self.goFollowingSlide = lambda:()   # do nothing
            self.activateFollowingDevice = lambda:self.controller.activatePrevDevice()
            prevDevice.gotoPrevSlide()
            self.lock.acquire()
            if not self.timerActive:
                self.timerActive = True
                self.gui.after(50, self.timerEvent)
            self.lock.release()
            return
        else:
            prevDevice.gotoPrevSlide()
            activeDevice.setBrightness(0)            
            prevDevice.setBrightness(self.controller.maxBrightness)
            self.controller.activatePrevDevice()
            return


    def startSlideshow(self):
        activeDevice = self.controller.activeDevice       
        
        if activeDevice == None:
            return
       
        self.fadeDelay = int(self.gui.fadeInput.get())
        
        self.slideshowDelay = int(self.gui.timerInput.get())

        activeDevice.setBrightness(100)
        self.state = 0
        self.slideshowActive = True
        self.slideshowPaused = False
        self.fadePaused = False
        self.lock.acquire()
        if not self.timerActive:
            self.gui.after(1000 * self.slideshowDelay, self.timerEvent)
            self.timerActive = True
        self.lock.release()
        
        return
   

    def pause(self):
        self.slideshowPaused = True
        self.fadePaused = True
        
            
    def resume(self):

        self.slideshowPaused = False
        self.fadePaused = False

        self.lock.acquire()
        if not self.timerActive:
            self.timerActive = True
            self.gui.after(50, self.timerEvent)
        self.lock.release()    


    def stopSlideshow(self):        
        
        self.state = 0
        self.slideshowActive = False        
        self.slidehowPaused = False
        self.fadePaused = False
        self.gui.pauseButton.config(text="pause")
        self.controller.resetDevices()


    def timerEvent(self):
        
        self.timerActive = False
        
        activeDevice = self.controller.activeDevice
        if activeDevice == None:
            return

        logger.debug("Timer Event: [" + self.states.get(self.state) + "]")

        

        if self.fadePaused:
            return
        
        #
        # IDLE
        #
        if self.state == 0:
            if self.slideshowActive and not self.slideshowPaused:
                self.nextSlide()
                self.gui.updateGUI()    
            return

        #
        # SINGLE_FADING_DOWN
        #
        if self.state == 1:
            self.timerCounter = self.timerCounter + 100
            level = int(0.2 * self.timerCounter / self.fadeDelay)
            if level < 100:
                activeDevice.setBrightness(100 - level)
                self.gui.updateGUI()
                self.lock.acquire()
                if not self.timerActive:                   
                    self.timerActive = True
                    self.gui.after(100, self.timerEvent)
                self.lock.release()
            else:
                activeDevice.setBrightness(0)
                self.goFollowingSlide()
                self.timerCounter = 0
                self.state = 2
                self.gui.updateGUI()
                self.lock.acquire()
                if not self.timerActive:
                    self.timerActive = True
                    self.gui.after(100, self.timerEvent)
                self.lock.release()
            return

        #
        # SINGLE_FADING_UP
        #
        if self.state == 2:
            self.timerCounter = self.timerCounter + 100
            level = int(0.2 * self.timerCounter / (self.fadeDelay + 1))

            if level < 100:
                activeDevice.setBrightness(level)
                self.gui.updateGUI()
                self.lock.acquire()
                if not self.timerActive:
                    self.timerActive = True
                    self.gui.after(100, self.timerEvent)
                self.lock.release()
            else:
                activeDevice.setBrightness(100)               
                self.timerCounter = 0                
                self.gui.updateGUI()
                if self.slideshowActive:
                    self.state = 0
                    self.lock.acquire()
                    if not self.timerActive:
                        self.timerActive = True
                        self.gui.after(1000 * self.slideshowDelay, self.timerEvent)
                    self.lock.release()    
            return

        #
        # DUAL_FADE
        #
        if self.state == 3:
            self.timerCounter = self.timerCounter + 100
            level = int(0.1 * self.timerCounter / (self.fadeDelay + 1))

            if level < 100:
                activeDevice.setBrightness(100 - level)
                self.followingDevice.setBrightness(level)
                self.gui.updateGUI()
                self.lock.acquire()
                if not self.timerActive:
                    self.timerActive = True
                    self.gui.after(100, self.timerEvent)
                self.lock.release()
            else:
                activeDevice.setBrightness(0)
                self.followingDevice.setBrightness(100)
                self.goFollowingSlide()
                self.activateFollowingDevice()
                self.timerCounter = 0
                self.state = 0
                self.gui.updateGUI()
                if self.slideshowActive:
                    self.lock.acquire()
                    if not self.timerActive:
                        self.timerActive = True
                        self.gui.after(1000 * self.slideshowDelay, self.timerEvent)
                    self.lock.release()
            return
       
            
        

    #
    # Helper
    #

    def isSingleProjector(self):
        return True if len(self.controller.devices) < 2 else False




            
class EktaproDevice:
    """
    Encapsulates the logic to control a single
    Ektapro slide projector.
    """

    def __init__(self, deviceInfo, serialDevice, internalID=0):
        if deviceInfo == None or len(deviceInfo) == 0 \
            or not (ord(deviceInfo[0]) % 8 == 6) \
            or not (ord(deviceInfo[1]) / 16 == 13) \
            or not (ord(deviceInfo[1]) % 2 == 0):            
                raise IOError, "invalid device"

        # from info string delivered by device 
        self.projektorID = ord(deviceInfo[0]) / 16
        self.projektorType = ord(deviceInfo[2]) / 16                             
        self.projektorVersion = str(ord(deviceInfo[2]) % 16) + "." \
                                + str(ord(deviceInfo[3]) / 16) \
                                + str(ord(deviceInfo[3]) % 16)

        self.powerFrequency = ord(deviceInfo[4]) & 128
        self.autoFocus = ord(deviceInfo[4]) & 64
        self.autoZero = ord(deviceInfo[4]) & 32
        self.lowLamp = ord(deviceInfo[4]) & 16
        self.traySize = 140 if ord(deviceInfo[4]) & 8 == 1 else 80
        self.activeLamp = ord(deviceInfo[4]) & 4
        self.standby = ord(deviceInfo[4]) & 2
        self.highLight = ord(deviceInfo[4]) & 1

        self.serialDevice = serialDevice

        # own temporary values
        self.brightness = 0        
        self.slide = 0

        self.internalID = internalID


    def __str__(self):
        modelStrings = {
            7: "4010 / 7000",
            4: "4020",
            5: "5000",
            6: "5020",
            8: "7010 / 7020",
            9: "9000",
            10: "9010 / 9020"
            }
            
        
        return "Kodak Ektapro " \
               + modelStrings.get(self.projektorType, "Unknown") \
               + " id=" + `self.projektorID` \
               + " Version " + `self.projektorVersion`

    def getDetails(self):
        return "Power frequency: " + ("60Hz" if self.powerFrequency == 1 else "50Hz") \
               + " Autofocus: " + ("On" if self.autoFocus == 1 else "Off") \
               + " Autozero: " + ("On" if self.autoZero == 1 else "Off") \
               + " Low lamp mode: " + ("On" if self.lowLamp == 1 else "Off") \
               + " Tray size: " + str(self.traySize) \
               + " Active lamp: " + ("L2" if self.activeLamp == 1 else "L1") \
               + " Standby: " + ("On" if self.standby == 1 else "Off") \
               + " High light: " + ("On" if self.highLight == 1 else "Off")


    def setStandby(self, on):
        c = EktaproCommand(self.projektorID).setStandby(on)
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())

    def setBrightness(self, brightness):
        c = EktaproCommand(self.projektorID).paramSetBrightness(brightness * 10)
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        self.brightness = brightness

    def resetSystem(self):
        c = EktaproCommand(self.projektorID).directResetSystem() 
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())

    def gotoSlide(self, slide):
        busy = True
        while busy:
            status = self.getSystemStatus()
            busy = (status["projector_status"] == 1)
            if busy:
                time.sleep(1)
        
        c = EktaproCommand(self.projektorID).paramRandomAccess(slide) 
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        self.slide = slide


    def gotoNextSlide(self):
        busy = True
        while busy:
            status = self.getSystemStatus()
            busy = (status["projector_status"] == 1)
            if busy:
                time.sleep(1)
        c = EktaproCommand(self.projektorID).directSlideForward()
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        
        self.slide = self.slide + 1
        
        if self.slide > self.traySize:
            self.slide = 0


    def gotoPrevSlide(self):
        busy = True
        while busy:
            status = self.getSystemStatus()
            busy = (status["projector_status"] == 1)
            if busy:
                time.sleep(1)
        c = EktaproCommand(self.projektorID).directSlideBackward()
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        
        self.slide = self.slide - 1
        if self.slide == -1:
            self.slide = self.traySize

    def getSystemStatus(self):
        c = EktaproCommand(self.projektorID).statusSystemStatus()
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        s = self.serialDevice.read(3)
        if not (ord(s[0]) % 8 == 6) \
           or not (ord(s[1]) / 16 == 12) \
           or not (ord(s[2]) % 4 == 3):            
            raise IOError, "invalid request response"
        status = {}
        status.update({"projector_id" : ord(s[0]) / 8})
        
        status.update({"lamp1_status" : ord(s[1]) & 8})
        status.update({"lamp2_status" : ord(s[1]) & 4})
        status.update({"projector_status" : ord(s[1]) & 2})
        status.update({"zero_position" : ord(s[1]) & 1})

        status.update({"slide_lift_motor_error" : ord(s[2]) & 128})
        status.update({"tray_transport_motor_error" : ord(s[2]) & 64})
        status.update({"command_error" : ord(s[2]) & 32})
        status.update({"overrun_error" : ord(s[2]) & 16})
        status.update({"buffer_overflow_error" : ord(s[2]) & 8})
        status.update({"framing_error" : ord(s[2]) & 4})
        return status

    def sync(self):
        c = EktaproCommand(self.projektorID).statusGetTrayPosition()
        logger.info("[" + str(self.internalID) + "] " + str(c))
        self.serialDevice.write(c.toData())
        s = self.serialDevice.read(3)
        if not (ord(s[0]) % 8 == 6) \
           or not (ord(s[1]) / 16 == 10):            
            raise IOError, "invalid request response"            
        self.slide = int(str(ord(s[2])))
    

class EktaproCommand:
    """
    Represents a single low level 3 byte command
    that is sent to an Ektapro slide projector
    by the user software.
    
    Permits easy construction of the 3 byte commands
    and decoding of 3 byte hex sequences to into
    a human readable string (for example for debugging).  
    """
    
    def __init__(self, *args):
        if len(args) == 1 :
            self.projektorID = args[0]
            self.initalized = False

        elif len(args) == 3:
            self.projektorID = args[0] / 8
            self.mode = args[0] % 8 / 2
            self.arg1 = args[1]
            self.arg2 = args[2]        
            self.initalized = True

        else:
            raise Exception, "argument count invalid"
        
    def toData(self):
        if not self.initalized:
            raise Exception, "Command not initialized"
        
        return chr(self.projektorID * 8 + self.mode * 2 + 1) \
               + chr(self.arg1) + chr(self.arg2)


    ###################################
    # Command  construction
    ###################################

    # Parameter mode

    def constructParameterCommand(self, command, param):
        self.mode = 0
        self.arg1 = command * 16 + param / 128 * 2
        self.arg2 = param % 128 * 2
        self.initalized = True
        

    def paramRandomAccess(self, slide):
        self.constructParameterCommand(0, slide)
        return self

    def paramSetBrightness(self, brightness):
        self.constructParameterCommand(1, brightness)
        return self

    def paramGroupAddress(self, group):
        self.constructParameterCommand(3, group)
        return self
        
    def paramFadeUp(self, time):
        self.constructParameterCommand(6, time + 128)
        return self

    def paramFadeDown(self, time):
        self.constructParameterCommand(6, time)
        return self
            
    def paramSetLowerLimitFading(self, time):
        self.constructParameterCommand(7, time)
        return self

    def paramSetUpperLimitFading(self, time):
        self.constructParameterCommand(8, time)
        return self

    # Set/Reset mode

    def constructSetResetCommand(self, option, on):
        self.mode = 1
        self.arg1 = option * 4 + (2 if on == True else 0)
        self.arg2 = 0
        self.initalized = True
        return self
        
    def setAutoFocus(self, on):
        self.constructSetResetCommand(0, on)
        return self
    
    def setHighlight(self, on):
        self.constructSetResetCommand(1, on)
        return self

    def setAutoShutter(self, on):
        self.constructSetResetCommand(3, on)
        return self

    def setBlockKeys(self, on):
        self.constructSetResetCommand(5, on)
        return self

    def setBlockFocus(self, on):
        self.constructSetResetCommand(2, on)
        return self

    def setStandby(self, on):
        self.constructSetResetCommand(7, on)
        return self
    
            
    # Direct mode

    def constructDirectModeCommand(self, command):
        self.mode = 2
        self.arg1 = command * 4
        self.arg2 = 0
        self.initalized = True

    def directSlideForward(self):
        self.constructDirectModeCommand(0)
        return self
    
    def directSlideBackward(self):
        self.constructDirectModeCommand(1)
        return self

    def directFocusForward(self):
        self.constructDirectModeCommand(2)
        return self

    def directFocusBackward(self):
        self.constructDirectModeCommand(3)
        return self

    def directFocusStop(self):
        self.constructDirectModeCommand(4)
        return self

    def directShutterOpen(self):
        self.constructDirectModeCommand(7)
        return self

    def directShutterClose(self):
        self.constructDirectModeCommand(8)
        return self

    def directResetSystem(self):
        self.constructDirectModeCommand(11)
        return self

    def directSwitchLamp(self):
        self.constructDirectModeCommand(12)
        return self

    def directClearErrorFlag(self):
        self.constructDirectModeCommand(13)
        return self

    def directStopFading(self):
        self.constructDirectModeCommand(15)
        return self

    # Status request mode

    def constructStatusRequestCommand(self, request):
        self.mode = 3
        self.arg1 = request * 16
        self.arg2 = 0
        self.initalized = True

    def statusGetTrayPosition(self):
        self.constructStatusRequestCommand(10)
        return self

    def statusGetKeys(self):
        self.constructStatusRequestCommand(11)
        return self
        
    def statusSystemStatus(self):
        self.constructStatusRequestCommand(12)
        return self
        
    def statusSystemReturn(self):
        self.constructStatusRequestCommand(13)
        return self
    

    ###################################
    # String  conversion
    ###################################
    
    def __str__(self):
        commandstring = {
            0: "Parameter Mode - " + self.parameterModeToString(),
            1: "Set/Reset Mode - " + self.setResetModeToString(),
            2: "Direct Mode - " + self.directModeToString(),
            3: "Status Request Mode - " + self.statusRequestToString()
            }

        
        return "Projektor " + str(self.projektorID) + " - " \
               + commandstring.get(self.mode, "Unknown Mode")
    

    def parameterModeToString(self):
        upDown = {
            0: "Down",
            1: "Up"
            }
        
        parametersettings = {
            0: "Random Access - Slide " + str(self.arg1 % 16 * 64 + self.arg2 / 2),
            1: "SetBrightness - " + str(self.arg1 % 16 * 64 + self.arg2 / 2),
            3: "Group Address - " + str(self.arg2 / 2),
            6: "Fade up/down - " + upDown.get(self.arg1 % 16 / 2, "?") + " - " \
                + str(self.arg2 / 2),
            7: "SetLowerLimit for Fading - " + str(self.arg1 % 16 * 64 + self.arg2 / 2),
            8: "SetUpperLimit for Fading - " + str(self.arg1 % 16 * 64 + self.arg2 / 2)
        }

        return parametersettings.get(self.arg1 / 16, "Unknown parameter")


    def setResetModeToString(self):
        setresetstring = {
            0: "AutoFocus on/off - ",
            1: "Highlight on/off - ",
            3: "AutoShutter on/off - ",
            5: "BlockKeys on/off - ",
            2: "BlockFocus on/off - ",
            7: "Standby on/off - "
            }

        onOff = {
            0: "Reset (off)",
            2: "Set (on)"
            }

        return setresetstring.get(self.arg1 / 4, "Unknown command") \
               + onOff.get(self.arg1 % 4, "?")


    def directModeToString(self):
        directModeString = {
            0: "Slide forward",
            1: "Slide backward",
            2: "Focus forward",
            3: "Focus backward",
            4: "Focus stop",
            7: "Shutter open",
            8: "Shutter close",
            11: "Reset system",
            12: "Switch lamp",
            13: "Clear error flags",
            15: "Stop fading"
            }

        if self.arg1 / 128 == 1:
            return "Direct User Mode"
        
        return directModeString.get(self.arg1 / 4, "Unknown command")


    def statusRequestToString(self):
        statusRequests = {
            10: "GetTray position",
            11: "GetKeys",
            12: "System status",
            13: "System return"
            }

        return statusRequests.get(self.arg1 / 16, "Unknown request")



class EktaproGUI(Tk):
    """
    Constructs the main program window
    and interfaces with the EktaproController
    and the TimerController to access the slide
    projectors.  
    """
    
    def __init__(self):
        self.controller = EktaproController()
        self.controller.initDevices()
      
        Tk.__init__(self)
        self.protocol('WM_DELETE_WINDOW', self.onQuit)
        self.wm_title("EktaproGUI")

        self.bind("<Prior>", self.priorPressed)
        self.bind("<Next>", self.nextPressed)
        
      
        self.brightness = 0
        self.slide = 1
        self.timerController = TimerController(self.controller, self)


        self.controlPanel = Frame(self)
        self.manualPanel = Frame(self)

        
        self.projektorList = Listbox(self, selectmode=SINGLE)

        for i in range(len(self.controller.devices)):            
            self.projektorList.insert(END, \
                                  "[" + str(i) + "] " + str(self.controller.devices[i]))

               
        if self.projektorList.size >= 1:          
            self.projektorList.selection_set(0)
            
        self.projektorList.bind("<ButtonRelease>", \
                                self.projektorSelectionChanged)
        self.projektorList.config(width=50)

        self.initButton = Button(self.controlPanel, \
                                 text="init", \
                                 command=self.initButtonPressed)
        self.nextButton = Button(self.controlPanel, \
                                 text="next slide", \
                                 command=self.nextSlidePressed)
        self.nextButton.config(state=DISABLED)
        self.prevButton = Button(self.controlPanel, \
                                 text="previous slide", \
                                 command=self.prevSlidePressed)
        self.prevButton.config(state=DISABLED)

        self.startButton = Button(self.controlPanel, \
                                  text="start timer", \
                                  command=self.startTimer)
        self.startButton.config(state=DISABLED)
        self.pauseButton = Button(self.controlPanel, \
                                  text="pause", \
                                  command=self.pauseTimer)        
        self.stopButton = Button(self.controlPanel, \
                                  text="stop", \
                                  command=self.stopTimer)
        self.stopButton.config(state=DISABLED)
        self.timerLabel = Label(self.controlPanel, \
                                text="delay:")        
        self.timerInput = Entry(self.controlPanel, \
                                width=3)
        self.timerInput.insert(0, "5")        
        self.timerInput.config(state=DISABLED)
        self.timerInput.bind("<KeyPress-Return>", self.inputValuesChanged)
        self.timerInput.bind("<ButtonRelease>", self.updateGUI)


        
        self.fadeLabel = Label(self.controlPanel, \
                                text="fade:")        
        self.fadeInput = Entry(self.controlPanel, \
                                width=3)
        self.fadeInput.insert(0, "1")
        self.fadeInput.config(state=DISABLED)
        self.fadeInput.bind("<KeyPress-Return>", self.inputValuesChanged)                        
        self.fadeInput.bind("<ButtonRelease>", self.updateGUI)
                         



        self.standbyButton = Button(self.controlPanel, \
                                    text="standby", \
                                    command=self.toggleStandby)
        self.standbyButton.config(state=DISABLED)
        self.syncButton = Button(self.controlPanel, \
                                 text="sync", \
                                 command=self.sync)
        self.syncButton.config(state=DISABLED)
        self.reconnectButton = Button(self.controlPanel, \
                                      text="reconnect", \
                                      command=self.reconnect)        
                                 

        self.cycle = IntVar()
        self.cycleButton = Checkbutton(self.controlPanel, \
                                       text="use all projectors", \
                                       variable=self.cycle, \
                                       command=self.cycleToggled)        

        self.brightnessScale = Scale(self.manualPanel, from_=0, to=100, resolution=1, \
                                     label="brightness")
        self.brightnessScale.set(self.brightness)
        self.brightnessScale.bind("<ButtonRelease>", self.brightnessChanged)
        self.brightnessScale.config(state=DISABLED)
        self.brightnessScale.config(orient=HORIZONTAL)
        self.brightnessScale.config(length=400)
        
    


        self.gotoSlideScale = Scale(self.manualPanel, \
                                    from_=0, to=self.controller.maxTray, \
                                    label="goto slide")
        self.gotoSlideScale.set(1)
        self.gotoSlideScale.bind("<ButtonRelease>", self.gotoSlideChanged)
        self.gotoSlideScale.config(state=DISABLED)
        self.gotoSlideScale.config(orient=HORIZONTAL)
        self.gotoSlideScale.config(length=400)
        
        

        self.controlPanel.pack(side=BOTTOM, anchor=W, fill=X)
        self.projektorList.pack(side=LEFT, fill=BOTH)
        self.manualPanel.pack(side=RIGHT, expand=1, fill=BOTH)
        
        self.initButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        
        self.prevButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.nextButton.pack(side=LEFT, anchor=N, padx=4, pady=4)        
        self.cycleButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.startButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.pauseButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.stopButton.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.timerLabel.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.timerInput.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.fadeLabel.pack(side=LEFT, anchor=N, padx=4, pady=4)
        self.fadeInput.pack(side=LEFT, anchor=N, padx=4, pady=4)
        
        

        
        self.syncButton.pack(side=RIGHT, anchor=N, padx=4, pady=4)
        self.standbyButton.pack(side=RIGHT, anchor=N, padx=4, pady=4)
        self.reconnectButton.pack(side=RIGHT, anchor=N, padx=4, pady=4)
        self.brightnessScale.pack(side=TOP, anchor=W, expand=1, fill=X)
        self.gotoSlideScale.pack(side=TOP, anchor=W , expand=1, fill=X)


        self.menubar = Menu(self)
        
        self.toolsmenu = Menu(self.menubar)
        self.helpmenu = Menu(self.menubar)
        self.filemenu = Menu(self.menubar)
         
        self.toolsmenu.add_command(label="Interpret HEX Sequence", \
                                   command=self.interpretHEXDialog)
       
        self.helpmenu.add_command(label="About EktaproGUI", \
                                  command=lambda:tkMessageBox.showinfo("About EktaproGUI", \
                                                                       "EktaproGUI 1.0 (C)opyright Julian Hoch 2010"))

        self.filemenu.add_command(label="Exit", command=self.onQuit)

        self.menubar.add_cascade(label="File", menu=self.filemenu)
        self.menubar.add_cascade(label="Tools", menu=self.toolsmenu)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)


        self.configure(menu=self.menubar)


    def initButtonPressed(self):
        self.controller.resetDevices()
        self.updateGUI()
        self.brightnessScale.config(state=NORMAL)
        self.gotoSlideScale.config(state=NORMAL)
        self.nextButton.config(state=NORMAL)
        self.prevButton.config(state=NORMAL)
        self.startButton.config(state=NORMAL)        
        self.timerInput.config(state=NORMAL)
        self.fadeInput.config(state=NORMAL)
        self.syncButton.config(state=NORMAL)
        self.standbyButton.config(state=NORMAL)


    def inputValuesChanged(self, event):        
        try:
            fadeDelay = int(self.fadeInput.get())
            slideshowDelay = int(self.timerInput.get())            
            if fadeDelay in range(0, 60):
                self.timerController.fadeDelay = fadeDelay
            if slideshowDelay in range(1, 60):                
                self.timerController.slideshowDelay = slideshowDelay            
        except Exception:
            pass
        self.updateGUI()
    

    def sync(self):
        self.controller.syncDevices()
        self.updateGUI()            


    def reconnect(self):
        self.controller.cleanUp()
        self.controller.initDevices()
        self.updateGUI()
        
        self.projektorList.delete(0, END)
        for i in range(len(self.controller.devices)):            
            self.projektorList.insert(END, \
                                  "[" + str(i) + "] " + str(self.controller.devices[i]))
               
        if self.projektorList.size >= 1:          
            self.projektorList.selection_set(0)


    def projektorSelectionChanged(self, event):
        items = map(int, self.projektorList.curselection())        
        if self.controller.setActiveDevice(items):
            self.updateGUI()


    def updateGUI(self, event=None):
        if self.controller.activeDevice == None:
            return
        
        self.brightness = self.controller.activeDevice.brightness
        self.brightnessScale.set(self.brightness)

        self.slide = self.controller.activeDevice.slide
        self.gotoSlideScale.set(self.slide)

        for i in range(self.projektorList.size()):
            if i == self.controller.activeIndex:
                self.projektorList.selection_set(i)
            else:
                self.projektorList.selection_clear(i)


    def brightnessChanged(self, event):
        newBrightness = self.brightnessScale.get()
        if not self.brightness == newBrightness \
           and not self.controller.activeDevice == None:
            self.controller.activeDevice.setBrightness(newBrightness)
            self.brightness = self.brightnessScale.get()


    def gotoSlideChanged(self, event):
        newSlide = self.gotoSlideScale.get()
        if not self.slide == newSlide:
            self.controller.activeDevice.gotoSlide(newSlide)
            self.slide = newSlide

  
    def nextSlidePressed(self):
        self.timerController.fadePaused = False
        self.timerController.nextSlide()
        self.updateGUI()

        
    def prevSlidePressed(self):
        self.timerController.fadePaused = False
        self.timerController.previousSlide()
        self.updateGUI()


    def startTimer(self):        
        self.stopButton.config(state=NORMAL)
        self.startButton.config(state=DISABLED)
        self.timerController.startSlideshow()        
            

    def pauseTimer(self):
        if self.timerController.fadePaused or self.timerController.slideshowPaused:
            self.pauseButton.config(text="pause")
            self.timerController.resume()
            self.updateGUI()            
        else:
            self.pauseButton.config(text="resume")
            self.timerController.pause()
            self.updateGUI()
        
        

    def stopTimer(self):        
        self.pauseButton.config(text="pause")
        self.stopButton.config(state=DISABLED)
        self.startButton.config(state=NORMAL)
        self.timerController.stopSlideshow()
        self.updateGUI()


    def cycleToggled(self):
        self.timerController.cycle = True if self.cycle.get() == 1 else False


    def interpretHEXDialog(self):        
        interpretDialog = InterpretHEXDialog(self) #@UnusedVariable


    def toggleStandby(self):
        if self.pauseButton.config()["text"][4] == "pause" \
           and self.pauseButton.config()["state"][4] == "normal":           
            self.pauseTimer()
        self.controller.toggleStandby()


    def nextPressed(self, event):
        if self.startButton.config()["state"][4] == "disabled":
            self.pauseTimer()            
        else:
            self.nextSlidePressed()
            

    def priorPressed(self, event):
        if self.startButton.config()["state"][4] == "disabled":        
            self.toggleStandby()
        else:      
            self.prevSlidePressed()


    def onQuit(self):
        self.controller.cleanUp()
        self.destroy()



class InterpretHEXDialog(tkSimpleDialog.Dialog):
    """
    A simple dialog that allows the user to
    enter 3 byte hex codes that represent
    ektapro commands and converts them to human
    readable strings.
    """

    def body(self, master):
        self.title("Interpret Ektapro HEX Command Sequence")

        self.description = Label(master, \
                                 text="Please enter 3 Byte HEX Code (for example: 03 1C 00):")
        self.description.pack(side=TOP, anchor=W)

        self.hexcommand = Entry(master, width=10)
        self.hexcommand.pack(side=TOP, anchor=W)

        return self.hexcommand

    def apply(self):
        inputString = self.hexcommand.get().replace(' ', '')
        
        if not len(inputString) == 6:
            tkMessageBox.showerror("Error", "Please enter exactly 3 Bytes (6 Characters)!")
            
        else:            
            b1 = int(inputString[0:2], 16)
            b2 = int(inputString[2:4], 16)
            b3 = int(inputString[4:6], 16)
            message = str(EktaproCommand(b1, b2, b3))
            tkMessageBox.showinfo("Command Result", \
                                   "Interpreted Command: " + message)
        
        
        

if __name__ == '__main__':    
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    mainWindow = EktaproGUI()
    mainWindow.mainloop()
