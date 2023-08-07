The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

- _RunScriptFunctions.ps1 - this contains a series of functions used by the RunScript*.* files to invoke specific actions. 
- RunScript*.CMD can be called via Task Scheduler and invoke the corresponding PS1 file. 
- The PS1 files load the functions from _RunScriptFunctions.ps1, then invoke the appropriate schedule. ~~

To use this: ~~~~
1. Edit the Invoke-DeltaRunProfiles function to build a series of run profiles to meet your needs. There are samples of checking pending imports and pending exports that can be re-used.
2. Update config.xml with thresholds, and email details
3. test thoroughly

