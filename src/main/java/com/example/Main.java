package com.example;
import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.time.Duration;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.util.logging.Logger;
import java.util.logging.Level;
import java.io.*;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.List;
import java.util.ArrayList;
import java.util.Set;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.HashMap;
import javax.swing.JOptionPane;
import javax.swing.JDialog;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Clipboard;
import java.awt.Toolkit;

public class Main {
    private static final Logger logger = Logger.getLogger(Main.class.getName());
    private static final String SELECT2_CHOICE_XPATH = "((//*[@class=\"select2-choice select2-default\"]))[1]";
    private static final String SAP_CONNECTION_NAME = "PRE Load Balance";
    private static final String SAP_SEND_VKEY_0 = "session.findById(\"wnd[0]\").sendVKey 0\n";
    private static final String TICKET_URL = "https://sdpondemand.manageengine.in/app/sandbox_60023490885_100725_iax/ui/requests";

    @SuppressWarnings({"java:S106", "java:S2629","java:S2142","java:S3776"})
    public static void main(String[] args) {
        logger.info("Starting WebDriver setup");
        WebDriverManager.chromedriver().setup();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("force-device-scale-factor=0.9");
        
        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(60));

        try {
            //open the Ticketing tool website
            logger.info("Starting Selenium automation");
            
            driver.get(TICKET_URL);
            
            // Login process
            long loginStartTime = System.currentTimeMillis();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Email address or mobile number']"))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@placeholder='Email address or mobile number']"))).sendKeys("ashishk@royalenfield.com");
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='nextbtn']"))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@autocomplete=\"username\"]))[1]"))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@type='text']"))).sendKeys("ashishk");
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@type='password']"))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@type='password']"))).sendKeys("Password@2003");
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='Sign in']"))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='auth-content-inner']"))).click();
            long loginEndTime = System.currentTimeMillis();
            double loginDuration = (loginEndTime - loginStartTime) / 1000.0;
            logger.log(Level.INFO, "Login process took {0} seconds", loginDuration);
            System.out.println("Login process took " + loginDuration + " seconds");

            boolean isFirstIteration = true;
            while (true) {
                try {
                    // Navigate to request screen to start/refresh the list
                    if (!isFirstIteration) {
                        driver.get(TICKET_URL);
                    }
                    isFirstIteration = false;

                    // Check if any tickets exist
                    if (!isTicketPresent(driver)) {
                        Thread.sleep(5000);
                        logger.info("No more tickets found. Exiting loop.");
                        JOptionPane pane = new JOptionPane("No more tickets available. All tickets processed.", JOptionPane.INFORMATION_MESSAGE);
                        JDialog dialog = pane.createDialog("Process Complete");
                        dialog.setAlwaysOnTop(true);
                        dialog.setVisible(true);
                        break;
                    }

                    List<String> extractedPORs = performSeleniumSteps(driver, wait);
                    
                    //Open SAP Logon and perform the process only if we have a valid POR extracted
                    if (extractedPORs != null && !extractedPORs.isEmpty()) {
                        openSAPLogon();
                        
                        // This will now process all PORs in a single SAP session for efficiency
                        Map<String, String> porSdnMap = performSAPProcessVBS("Bijus", "Royal@26", extractedPORs);

                        StringBuilder resolutionBuilder = new StringBuilder();
                        resolutionBuilder.append("HI, PLEASE FIND THE SDN NUMBER BELOW \n");
                        boolean valueFound = false; // Tracks if any value (SDN or message) was found
                        boolean allSdnsFound = !extractedPORs.isEmpty(); // Start with true only if there are PORs to check

                        // Iterate through the original list to maintain order and build the resolution string
                        for (String por : extractedPORs) {
                            // Get the corresponding SDN from the map we populated earlier
                            String extractedValue = porSdnMap.get(por);
                            
                            if (extractedValue != null && !extractedValue.isEmpty()) {
                                logger.log(Level.INFO, "Extracted Value for POR {0}: {1}", new Object[]{por, extractedValue});
                                if (extractedValue.matches("\\d+")) {
                                    resolutionBuilder.append("POR: ").append(por).append(" SDN: ").append(extractedValue).append("\n");
                                } else {
                                    resolutionBuilder.append("POR: ").append(por).append(" Message: ").append(extractedValue).append("\n");
                                    allSdnsFound = false; // A message means not all are SDNs
                                }
                                valueFound = true;
                            } else {
                                logger.log(Level.WARNING, "Could not find SDN or Message for POR {0}", por);
                                allSdnsFound = false; // A missing value means not all are SDNs
                            }
                        }

                        // Only proceed if at least one value (SDN or message) was extracted from SAP
                        if (valueFound) {
                            if (allSdnsFound) {
                                // All PORs have a valid SDN, go to RESOLUTION tab
                                String ticketID = performResolutionUpdate(driver, wait, resolutionBuilder.toString());
                                if (ticketID != null && !ticketID.isEmpty()) {
                                    saveTicketToLocalCSV(ticketID);
                                }
                            } else {
                                // At least one POR was missing an SDN, go to NOTES tab
                                addNoteAndSetInProgress(driver, wait, resolutionBuilder.toString());
                            }
                        } else {
                            logger.warning("Skipping ticket resolution because no value (SDN or Message) was extracted from SAP.");
                            System.out.println("Skipping ticket resolution because no value was extracted from SAP.");
                        }
                    } else {
                        logger.warning("Skipping SAP process and ticket resolution because POR was not extracted.");
                        System.out.println("Skipping SAP process and ticket resolution because POR was not extracted.");
                    }
                } catch (Exception e) {
                    logger.log(Level.SEVERE, "Error processing ticket in loop", e);
                }
            }

        } catch (Exception e) {
            logger.log(Level.SEVERE, "Critical failure in automation process", e);
        } finally {
            logger.info("Process finished. Browser remains open.");
        }
    }

    private static void addNoteAndSetInProgress(WebDriver driver, WebDriverWait wait, String noteText) throws InterruptedException {
        logger.info("SDN not found for all PORs. Adding a note and setting status to In Progress.");
        
        // 1. Click on note tab
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@sdtooltip=\"BOTTOMFAST\"]))[11]"))).click();
        
        // 2. Wait
        Thread.sleep(3000);
        
        // 3. Enter iframe
        logger.info("Switching to note frame...");
        wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.cssSelector("[class=\"ze_area\"]")));
        
        // 4. Enter text
        logger.info("Inside frame. Sending note text...");
        WebElement body = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"ze_body\"]")));
        body.click();
        body.sendKeys(noteText);
        
        // 5. Wait
        Thread.sleep(2000);
        
        // 6. Exit iframe
        logger.info("Switching back to default content...");
        driver.switchTo().defaultContent();
        
        // 7. Wait
        Thread.sleep(2000);
        
        // 8. Click save
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[text()='Save']))[3]"))).click();
        
        // 9. Wait
        Thread.sleep(3000);
        
        // 10. Click status dropdown
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"wo-result dropdown-arrow-after-ico spot-editable\"]))[1]"))).click();
        
        // 11. Click "In Progress"
        clickDropdownOptionWithScroll(driver, wait, 7);
        
        // Verify status
        try {
            new WebDriverWait(driver, Duration.ofSeconds(10)).until(ExpectedConditions.textToBePresentInElementLocated(By.cssSelector("[data-sdplayoutid=\"status\"]"), "In Progress"));
            logger.info("Status verified as In Progress.");
        } catch (Exception e) {
            logger.warning("Status verification for 'In Progress' failed: " + e.getMessage());
        }
    }

    private static boolean isTicketPresent(WebDriver driver) {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(20))
                .until(ExpectedConditions.presenceOfElementLocated(By.xpath("((//*[@class=\"listview-display-id\"]))[1]")));
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    // This method performs the Selenium steps to extract the POR and navigate through the ticketing system. It returns the extracted POR for use in the SAP process.
    @SuppressWarnings({"java:S106", "java:S2629","java:S2142"})
        private static List<String> performSeleniumSteps(WebDriver driver, WebDriverWait wait) throws InterruptedException {
        Thread.sleep(5000);
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"listview-display-id\"]))[1]"))).click();
        Thread.sleep(3000);

        List<String> extractedPORs = new ArrayList<>();
        try {
            logger.info("Attempting to find POR element...");
            //WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("((//*[@class=\"req-data details-inner-view\"]))[1]")));
            Thread.sleep(2000);
            WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[class=\"conversation-parent-container\"]")));
            String text = element.getText();
            logger.info("Raw text found in element: " + text);
            System.out.println("Raw text found in element: " + text);

            Pattern pattern = Pattern.compile("\\bPOR[A-Z0-9]{13}\\b");
            Matcher matcher = pattern.matcher(text);
            Set<String> uniquePORs = new LinkedHashSet<>();
            while (matcher.find()) {
                uniquePORs.add(matcher.group(0));
            }
            extractedPORs.addAll(uniquePORs);
            
            if (!extractedPORs.isEmpty()) {
                logger.log(Level.INFO, "Extracted PORs: {0}", extractedPORs);
                System.out.println("Extracted PORs: " + extractedPORs);
            } else {
                logger.warning("Regex did not match any POR in the text.");
                System.out.println("Regex did not match any POR in the text.");
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Could not extract POR: {0}", e.getMessage());
            System.out.println("Could not extract POR: " + e.getMessage());
        }

        // Provide the remark to the ticket 
        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"request-list-tech-name\"]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-arrow\"]))[4]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-result-label\"]))[1]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"active-button formstylebutton formstylebutton\"]"))).click();
        Thread.sleep(2000);
        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[tabid=\"requestDetails\"]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("id(\"details_inner_view\")/DIV[2]/DIV[1]/DIV[9]/DIV[1]/DIV[1]/H2[1]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("id(\"details_inner_view\")/DIV[2]/DIV[1]/DIV[9]/DIV[1]/DIV[2]/DIV[1]/DIV[2]/SPAN[1]/A[1]"))).click();
        for (int i : new int[]{1, 13, 5}) {
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath(SELECT2_CHOICE_XPATH))).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-result-label\"]))["+i+"]"))).click();
        }
        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"saveicon\"]"))).click();
        Thread.sleep(2000);
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("id(\"details_inner_view\")/DIV[2]/DIV[1]/DIV[9]/DIV[1]/DIV[3]/DIV[2]/DIV[2]/SPAN[1]/A[1]"))).click();
        Thread.sleep(5000);
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-arrow\"]))[2]"))).click();
        Thread.sleep(3000);
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-result-label\"]))[5]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"saveicon\"]"))).click();
        return extractedPORs;
    }

    // This method performs the resolution update steps in the ticketing system using Selenium. It uses the extracted SDN from the SAP process to update the ticket and then extracts the Ticket ID for saving.
    @SuppressWarnings({"java:S106", "java:S2629","java:S2142", "java:S3776"})
        private static String performResolutionUpdate(WebDriver driver, WebDriverWait wait, String resolutionText) throws InterruptedException {
            logger.info("Ready to update resolution.");
            
            // Ensure browser is focused
            driver.switchTo().window(driver.getWindowHandle());
            
            WebElement resolutionTab = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[title=\"Resolution\"]")));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", resolutionTab);
            
            logger.info("Switching to resolution frame...");
            wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.cssSelector("[class=\"ze_area\"]")));
            logger.info("Inside frame. Sending text...");
            WebElement body = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"ze_body\"]")));
            body.click();
            body.sendKeys(resolutionText);
            logger.info("Switching back to default content...");
            Thread.sleep(3000);
            driver.switchTo().defaultContent();
            Thread.sleep(3000);
            
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("((//*[@class=\"select2-arrow\"]))[3]"))).click();
            Thread.sleep(5000);
            
            clickDropdownOptionWithScroll(driver, wait, 40);
            Thread.sleep(7000);
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("[class=\"active-button formstylebutton form-action formstylebutton\"]"))).click();
            Thread.sleep(5000);

            try {
                wait.until(ExpectedConditions.textToBePresentInElementLocated(By.cssSelector("[data-sdplayoutid=\"status\"]"), "Resolved"));
                logger.info("Status verified as Resolved.");
            } catch (Exception e) {
                logger.warning("Status verification failed (might be case sensitivity or timeout): " + e.getMessage());
            }

            // Extract Ticket ID
            String ticketID = "";
            try {
                WebElement idElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[class=\"rhs-req-display-id\"]")));
                ticketID = idElement.getText();
                logger.info("Extracted Ticket ID: " + ticketID);
                System.out.println("Extracted Ticket ID: " + ticketID);
            } catch (Exception e) {
                logger.warning("Failed to extract Ticket ID: " + e.getMessage());
            }
            return ticketID;
        }

        // This method opens the SAP Logon application. 
        public static void openSAPLogon() {
            try {
                logger.info("Launching SAP GUI...");
                ProcessBuilder pb = new ProcessBuilder("C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe");
                pb.start();
                Thread.sleep(8000); 
            } catch (InterruptedException e) {
                logger.log(Level.SEVERE, "Interrupted while launching SAP GUI", e);
                Thread.currentThread().interrupt();
            } catch (Exception e) {
                logger.log(Level.SEVERE, "Could not start SAP EXE", e);
            }
        }
        
        // This method creates and executes a VBScript to automate the SAP GUI process. It logs in, navigates to the relevant table, extracts the SDN based on the provided POR, and returns it for use in the Selenium resolution update.
        @SuppressWarnings({"java:S106", "java:S2629","java:S2142","java:S3776"})
        public static Map<String, String> performSAPProcessVBS(String user, String pass, List<String> porNumbers) {
            Map<String, String> finalPorSdnMap = new HashMap<>();
            if (porNumbers == null || porNumbers.isEmpty()) {
                return finalPorSdnMap;
            }

            String vbsPath = "sap_bridge.vbs";
            String outputPath = "sap_output.txt";

            try {
                // Copy POR numbers to system clipboard for SAP to paste
                String clipboardData = String.join("\n", porNumbers);
                StringSelection stringSelection = new StringSelection(clipboardData);
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                clipboard.setContents(stringSelection, null);

                StringBuilder vbsContent = new StringBuilder();
                // --- VBScript Generation ---
                vbsContent.append("If Not IsObject(application) Then\n");
                vbsContent.append("   Set SapGuiAuto  = GetObject(\"SAPGUI\")\n");
                vbsContent.append("   Set application = SapGuiAuto.GetScriptingEngine\n");
                vbsContent.append("End If\n");
                vbsContent.append("Set connection = application.OpenConnection(\"").append(SAP_CONNECTION_NAME).append("\", True)\n");
                vbsContent.append("Set session = connection.Children(0)\n");
                // Maximize SAP window to ensure it is visible
                vbsContent.append("session.findById(\"wnd[0]\").maximize\n");
                
                // Login logic
                vbsContent.append("session.findById(\"wnd[0]/usr/txtRSYST-BNAME\").text = \"").append(user).append("\"\n");
                vbsContent.append("session.findById(\"wnd[0]/usr/pwdRSYST-BCODE\").text = \"").append(pass).append("\"\n");
                vbsContent.append(SAP_SEND_VKEY_0);

                vbsContent.append("session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"/nSE16\"\n"); 
                vbsContent.append(SAP_SEND_VKEY_0); 
                vbsContent.append("session.findById(\"wnd[0]/usr/ctxtDATABROWSE-TABLENAME\").text = \"ZSD_CDS_PO_IF\"\n"); 
                vbsContent.append(SAP_SEND_VKEY_0); 

                // Click 'Multiple Selection' button for the POR field
                vbsContent.append("session.findById(\"wnd[0]/usr/btn%_I4_%_APP_%-VALU_PUSH\").press\n"); 

                // Paste from clipboard (Shift+F12)
                vbsContent.append("session.findById(\"wnd[1]/tbar[0]/btn[24]\").press\n");

                // Execute multiple selection and then the main search
                vbsContent.append("session.findById(\"wnd[1]/tbar[0]/btn[8]\").press\n"); 
                vbsContent.append("session.findById(\"wnd[0]/tbar[1]/btn[8]\").press\n"); 
                
                vbsContent.append("WScript.Sleep 10000\n"); 

                vbsContent.append("Set grid = session.findById(\"wnd[0]/usr/cntlGRID1/shellcont/shell\")\n");

                // Sort by Document (SDN) column descending to bring the correct/latest record to the top
                vbsContent.append("grid.setCurrentCell -1, \"SDN\"\n");
                vbsContent.append("grid.selectColumn \"SDN\"\n");
                vbsContent.append("session.findById(\"wnd[0]/tbar[1]/btn[40]\").press\n"); // Descending sort button
                vbsContent.append("WScript.Sleep 2000\n"); // Wait for sort to apply

                // Write to file
                vbsContent.append("Set fso = CreateObject(\"Scripting.FileSystemObject\")\n");
                vbsContent.append("Set f = fso.CreateTextFile(\"").append(outputPath).append("\", True)\n");

                // Loop through all rows and extract data only if both POR and SDN are present
                vbsContent.append("rowCount = grid.RowCount\n");
                vbsContent.append("For i = 0 To rowCount - 1\n");
                vbsContent.append("  grid.firstVisibleRow = i\n");
                // The getCellValue method requires the technical field name, not the display label.
                vbsContent.append("  por_val = grid.getCellValue(i, \"DMSPONO\")\n"); // 'DMSPONO' is the technical name for the 'Customer Reference' column.
                vbsContent.append("  sdn_val = grid.getCellValue(i, \"SDN\")\n");   // 'SDN' is the technical name for the 'Document' column.
                
                vbsContent.append("  final_val = sdn_val\n");
                vbsContent.append("  If Trim(final_val) = \"\" Then\n");
                vbsContent.append("    On Error Resume Next\n");
                vbsContent.append("    final_val = grid.getCellValue(i, \"MESSAGE\")\n"); // Attempt to read MESSAGE column
                vbsContent.append("    If Err.Number <> 0 Then final_val = \"\" \n");
                vbsContent.append("    On Error GoTo 0\n");
                vbsContent.append("  End If\n");

                vbsContent.append("  If Trim(por_val) <> \"\" And Trim(final_val) <> \"\" Then\n");
                vbsContent.append("    f.WriteLine por_val & \",\" & final_val\n");
                vbsContent.append("  End If\n");
                vbsContent.append("Next\n");

                vbsContent.append("f.Close\n");
                vbsContent.append("session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"/nex\"\n");
                vbsContent.append(SAP_SEND_VKEY_0);

                // --- VBScript Execution ---
                Files.write(Paths.get(vbsPath), vbsContent.toString().getBytes(StandardCharsets.UTF_8));
                ProcessBuilder pb = new ProcessBuilder("wscript", vbsPath);
                Process p = pb.start();
                p.waitFor();

                // --- Java-side Processing ---
                // Read all POR,SDN pairs from the output file
                File outputFile = new File(outputPath);
                if (outputFile.exists()) {
                    List<String> lines = Files.readAllLines(Paths.get(outputPath), StandardCharsets.UTF_8);

                    for (String line : lines) {
                        String[] parts = line.split(",", 2);
                        if (parts.length == 2) {
                            String por = parts[0].trim();
                            String sdn = parts[1].trim();
                            if (!por.isEmpty() && !sdn.isEmpty()) {
                                // Per user, one POR has one unique SDN. Just map it directly.
                                finalPorSdnMap.put(por, sdn);
                            }
                        }
                    }
                }

                Files.deleteIfExists(Paths.get(vbsPath));
                Files.deleteIfExists(Paths.get(outputPath));

            } catch (InterruptedException e) {
                logger.log(Level.SEVERE, "Thread interrupted during SAP VBScript execution", e);
                Thread.currentThread().interrupt();
            } catch (Exception e) {
                logger.log(Level.SEVERE, "SAP VBScript Error", e);
            }
            return finalPorSdnMap;
        }

    // This method is a helper to click on dropdown options that may require scrolling. It uses JavaScript to scroll the element into view before clicking, and includes error handling to log any issues encountered during the process.
    @SuppressWarnings({"java:S106", "java:S2629","java:S2142","java:S3457"})
        private static void clickDropdownOptionWithScroll(WebDriver driver, WebDriverWait wait, int index) {
            String xpath = "((//*[@class=\"select2-result-label\"]))[" + index + "]";
            try {
                // Wait for the element to be present in the DOM
                WebElement element = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(xpath)));
                
                // Scroll into view using JavaScript
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
                Thread.sleep(1000); 
                
                // Click
                wait.until(ExpectedConditions.elementToBeClickable(element)).click();
            } catch (Exception e) {
                logger.log(Level.WARNING, "Failed to scroll/click index " + index, e);
            }
        }

    // This method saves the extracted Ticket ID to a local CSV file.    
    @SuppressWarnings({"java:S106", "java:S2629","java:S2142"})    
        private static void saveTicketToLocalCSV(String ticketID) {
            try {
                String filePath = "ExtractedTickets.csv";
                String content = ticketID + "\n";
                Files.write(Paths.get(filePath), content.getBytes(StandardCharsets.UTF_8), StandardOpenOption.CREATE, StandardOpenOption.APPEND);
                
                logger.info("Ticket ID saved to local CSV: " + filePath);
                System.out.println("Ticket ID saved to local CSV: " + filePath);

            } catch (java.nio.file.FileSystemException e) {
                String errorMessage = "Failed to write to ExtractedTickets.csv. Please ensure the file is not open in another program (like Excel) and try again.";
                logger.log(Level.SEVERE, errorMessage, e);
                System.out.println("ERROR: " + errorMessage);

            } catch (Exception e) {
                logger.log(Level.SEVERE, "Failed to save Ticket ID to CSV", e);
            }
        }
}