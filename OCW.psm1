<#
    This module is used to calculate the best Optimum Charge Weight for Reloading ammunition.  All methods are based on Dan Newberry's Optimum Charge Weight Load Development

    See the following links

    http://optimalchargeweight.embarqspace.com/

    http://www.bangsteel.com/
#>

function Get-OCWGroupTriangulation {

<#
    .Synopsis
        Triangulate Shot Groupings

    .Description
        Calculates the groups center, Distance and direction from bullseye.  These are known as the Point of Impact (POI).

    .Parameter Shot
        X,Y in inches of the shot on target.

    .Parameter Name
        Name of the group.
#>
    [CmdletBinding()]
    Param ( 
        [Parameter (Mandatory = $True)]
        [String]$Name,

        [Parameter (Mandatory = $True, ValueFromPipeline = $True) ]
        [PSObject[]]$Shot
    )

    Begin {
        $ShotsInGroup =@()
    }

    Process {
        # ----- Collect Shots into object
        Foreach ( $S in $Shot ) {
            Write-verbose "Adding Shot to $Name"
            $Hole = New-Object -TypeName PSObject -Property (@{
                'X' = $S.X
                'Y' = $S.Y
            })

            $ShotsInGroup += $Hole
        }

    }

    End {
        $Group = New-Object -TypeName PSObject -Property (@{
            'Name' = $Name
            'Shots' = $ShotsInGroup    
        })

        Write-Verbose "Perform Calculations"
        switch ( (Measure-object $Group.Shots).Count ) {
            2 {
                Write-Verbose "Group of 2"
                $POI = New-Object -TypeName PSObject -Property (@{
                    'X' = ($Group.Shots[0].X + $Group.Shots[1].X)/2
                    'Y' = ($Group.Shots[0].Y + $Group.Shots[1].Y)/2
                })
            }

            3 {
                Write-verbose "Group of 3"
                $POI = New-Object -TypeName PSObject -Property (@{
                    'X' = ($Group.Shots[0].X + $Group.Shots[1].X + $Group.Shots[2].X)/3
                    'Y' = ($Group.Shots[0].Y + $Group.Shots[1].Y + $Group.Shots[2].Y)/3
                })
            }

            Default {
                Throw "Get-OCWGroupTriagulation : Unknown number of shots in group"
            }
        }

        $Group | Add-Member -MemberType NoteProperty -Name POI -Value $POI
    }
}
