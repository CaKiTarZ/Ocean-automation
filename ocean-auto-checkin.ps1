#
#----------------------[Variables]---------------------------------------------------------------#
#
#Set Username and Pass
$user = ""
$pass = ''

# uncomment for testing
#$action = "checkout"

#
#----------------------[Parameters]---------------------------------------------------------------#
#
param (
    
    [Parameter(Mandatory=$true)]  
    [String] $action

) 

if (($action -ne "checkin") -and ($action -ne "checkout")){
    $errorMessage = "Incorrect action set, value must be 'checkin' or 'checkout'."
    throw $errorMessage
}

#
#----------------------[Init]--------------------------------------------------------------------#
#
# check if module PoShKeePass is installed
#if (!(Import-Module -Name "PoShKeePass")) {
    # module missing, proceed to install
#    Install-Module -Name "PoShKeePass" -Force
#}

# import module
#Import-Module -Name "PoShKeePass" -Force

# working directory
$workingPath = Get-Location

# add the working directory to the env path
# in order for chromedriver to work
if (($env:Path -split ";") -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}

# check the Path environment variable
$env:Path -split ';'

# add the webdrivell.dll to powershell
Add-Type -Path "$($workingPath)\WebDriver.dll"


#
#-------------------------[Execution]------------------------------------------------------------#
#
# create a new ChromeDriver Object instance.
$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver

# Launch a browser and go to URL
$ChromeDriver.Navigate().GoToURL("https://ocean.mediapro.tv/")

Start-Sleep -Seconds 2s

# Login Corporativo Button XPath: //*[@id="ng-app"]/body/div/div/div/div/div[5]/div[1]/form/div[4]/a
# Click Loggin Corporativo
$ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div/div/div/div/div[5]/div[1]/form/div[4]/a').Click()

Start-Sleep -Seconds 2s

# Mediapro User Box XPath: //*[@id="i0116"]
# Enter Username 
$ChromeDriver.FindElementByXPath('//*[@id="i0116"]').SendKeys($user)

# MS login Next button XPath: //*[@id="idSIButton9"]
# Click on Next
$ChromeDriver.FindElementByXPath('//*[@id="idSIButton9"]').Click()

Start-Sleep -Seconds 2s

# MS password Box XPath: //*[@id="i0118"]
# Enter Password
$ChromeDriver.FindElementByXPath('//*[@id="i0118"]').SendKeys($pass)

# MS login button XPath: //*[@id="idSIButton9"]
# Click on Login button
$ChromeDriver.FindElementByXPath('//*[@id="idSIButton9"]').Click()

Start-Sleep -Seconds 2s

# If message to mantain session appears:
# "¿Quiere mantener la sesión iniciada?"
# XPath: /html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[1]

$mantainSession = $ChromeDriver.FindElementByXPath('/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[1]')
if ($mantainSession) {
    if ($mantainSession.Text -eq "¿Quiere mantener la sesión iniciada?") {
        # the box appeared we have to mark "no volver a mostar"
        # XPath: html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[1]/div/label/input
        $ChromeDriver.FindElementByXPath('html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[1]/div/label/input').Click()
        
        # then click "sí" button
        # si button XPath: //*[@id="idSIButton9"]
        $ChromeDriver.FindElementByXPath('//*[@id="idSIButton9"]').Click()

    }
}

Start-Sleep -Seconds 2s

# check if we are inside ocean portal
# marcajes button XPath: //*[@id="ng-app"]/body/div/div/div[2]/div[1]/div/div/div/div/div/div/a/i
$marcajes = $ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div/div/div[2]/div[1]/div/div/div/div/div/div/a')
if ($marcajes) {
    if ($marcajes.Text -eq "Marcajes") {
        # we are in
        
        # show menu
        # menu XPath: //*[@id="menuboton"]/i
        #$ChromeDriver.FindElementByXPath('//*[@id="menuboton"]/i').Click()
       
        # proceed to "Realizar marcaje Manual"
        # Marcaje manual XPath: //*[@id="ng-app"]/body/div/div/div[1]/div[2]/div/div/a/span
        $ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div/div/div[1]/div[2]/div/div/a/span').Click()

        # Fichaje entrada o salida
        # Marcaje sin incidencia XPath: //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[1]/div[1]/ul/li[1]/a
        $ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[1]/div[1]/ul/li[1]/a').Click()
        
        # Fichaje entrada o salida comida:
        # COMIDA (001) XPath: //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[1]/div[1]/ul/li[2]/a
        #$ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[1]/div[1]/ul/li[2]/a').Click()
        Start-Sleep -Seconds 2s

        if ($action -eq "checkout") {
            # Marcaje de Salida
            # Marcaje de Salida XPath: //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[2]/div/a/span
            # //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[2]/div/a/i
            # //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[2]/div/a/span
            $ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[2]/div/a/span').Click()
        }

        if ($action -eq "checkin") {
            # Marcaje de Entrada
            # Macaje de Entrada XPath: //*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[1]/div/a/span
            #//*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[1]/div/a/i
            $ChromeDriver.FindElementByXPath('//*[@id="ng-app"]/body/div[3]/div/div/div/div[2]/div/div/div/div/div[4]/div[1]/div[2]/div/div[1]/div/a/i').Click()
        }
        
        # refresh marcajes page
        $ChromeDriver.Navigate().GoToURL("https://ocean.mediapro.tv/#/marcajes")
        
        Start-Sleep -Seconds 2s
        
        # scroll down 500 pixel to get clear view
        $ChromeDriver.ExecuteScript("window.scrollBy(0,200)")
        
        # take screenshot
        $screenshotBase64 = $ChromeDriver.GetScreenshot()
        $Image = [Drawing.Bitmap]::FromStream([IO.MemoryStream][Convert]::FromBase64String($screenshotBase64))
        $date = Get-Date -Format FileDateTime
        $Image.Save("$($action)_$($date).jpg")
        
        # To do sent telegram or email with screenshot
    }
}

# Cleanup
$ChromeDriver.Close()
$ChromeDriver.Quit()