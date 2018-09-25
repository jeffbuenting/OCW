$ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path

$ModuleName = $ModulePath | Split-Path -Leaf

# ----- Remove and then import the module.  This is so any new changes are imported.
Get-Module -Name $ModuleName -All | Remove-Module -Force -Verbose

Import-Module "$ModulePath\$ModuleName.PSD1" -Force -ErrorAction Stop -Scope Global -Verbose

#-------------------------------------------------------------------------------------

Write-Output "`n`n"

Describe "OCW : Get-Get-OCWGroupTriangulation" {
    # ----- Get Function Help
    # ----- Pester to test Comment based help
    # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html
    Context "Help" {

        $H = Help Get-OCWGroupTriangulation -Full

        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples
        it "Example - Count should be greater than 0"{
            $H.examples.example.code.count | Should BeGreaterthan 0
        }
            
        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 

    Context Execution {

        It "Should Error if an unknown number of shots are passed" {
            { @{ 'X'=1.5;'Y'=.2},@{ 'X'=.5;'Y'=.5},@{ 'X'=.75;'Y'=1} | Get-OCWGroupTriangulation -Name Test -verbose } | Should Throw
        }

    }
  
    Context Output {
        It "Should return Custom Object" {
            Get-OCWGroupTriangulation -Name Group1 -Shot Shot1 | Should BeOfType 'PSObject'
        }
    }
}

#-------------------------------------------------------------------------------------

Write-Output "`n`n"

Describe "OCW : Optimize-OCWLoad" {
    # ----- Get Function Help
    # ----- Pester to test Comment based help
    # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html
    Context "Help" {

        $H = Help Optimize-OCWLoad -Full

        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples
        it "Example - Count should be greater than 0"{
            $H.examples.example.code.count | Should BeGreaterthan 0
        }
            
        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 

    
}