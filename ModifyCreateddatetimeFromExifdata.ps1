$ErrorActionPreference = "stop"
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.IO")

# 撮影日時取得処理
# $param1 : Bitmap object
# return  : $propDate
function GetExifDatetime($sourceImage) {
    $propDate = ""

    # Exif 情報取得
    $private:properties = $sourceImage.PropertyItems
    foreach ($property in $properties) {
        if($property.Id -eq 306) {
            #0x0132 PropertyTagDateTime
            $byteAry = $property.Value
            $byteAry[4] = 47
            $byteAry[7] = 47
            $propDate = [System.Text.Encoding]::ASCII.GetString($byteAry)
            return $propDate
            echo "PropertyTagDateTime:$propDate"
        }
        elseif($property.Id -eq 36867) {
            #0x9003 PropertyTagExifDTOrig
            $byteAry = $property.Value
            $byteAry[4] = 47
            $byteAry[7] = 47
            $propDate = [System.Text.Encoding]::ASCII.GetString($byteAry)
            return $propDate
            echo "PropertyTagExifDTOrig:$propDate"
        }
        elseif($property.Id -eq 36868) {
            #0x9004 PropertyTagExifDTDigitized
            $byteAry = $property.Value
            $byteAry[4] = 47
            $byteAry[7] = 47
            $propDate = [System.Text.Encoding]::ASCII.GetString($byteAry)
            return $propDate
            echo "PropertyTagExifDTDigitized:$propDate "
        }
    }

    return $propDate
}

# Main
# param1 : Original fullpath file name
# return : none
function script:Main($originalFileName) {
    $private:sourceImage = $null

    try {
        # 元画像取得(Bitmap)
        $sourceImage = New-Object System.Drawing.Bitmap($originalFileName)

        # 撮影日取得
        $i = GetExifDatetime $sourceImage

        # 元Bitmapを閉じる
        $sourceImage.Dispose()

        echo "終了しました:$originalFileName"
    }
    catch {
        echo "例外が発生しました:$error"
    }
    finally {
        try {
            if($null -ne $sourceImage) {
                $sourceImage.Dispose()
            }
        }
        catch {
            echo "画像クローズ時の例外:$error"
        }
    }

    Set-ItemProperty $originalFileName -Name CreationTime -Value $i
    Set-ItemProperty $originalFileName -Name LastWriteTime -Value $i
}

# 処理実行
$d = Get-ChildItem
$fullpath = (Get-ChildItem -Recurse -File $d).FullName
foreach($filename in $fullpath) {
    Main $filename
}

