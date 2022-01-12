//go:build windows
// +build windows

package main

import (
	"fmt"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	ole.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("Outlook.Application")
	outlook, _ := unknown.QueryInterface(ole.IID_IDispatch)
	//oleutil.PutProperty(outlook, "Visible", true)
	ns := oleutil.MustCallMethod(outlook, "GetNamespace", "MAPI").ToIDispatch()
	folder := oleutil.MustCallMethod(ns, "GetDefaultFolder", 10).ToIDispatch()
	contacts := oleutil.MustCallMethod(folder, "Items").ToIDispatch()
	count := oleutil.MustGetProperty(contacts, "Count").Value().(int32)
	for i := 1; i <= int(count); i++ {
		item, err := oleutil.GetProperty(contacts, "Item", i)
		if err == nil && item.VT == ole.VT_DISPATCH {
			if value, err := oleutil.GetProperty(item.ToIDispatch(), "FullName"); err == nil {
				fmt.Println(value.Value())
			}
		}
	}

	omsgSub := "Hello"
	omsgName := "murali.achanta@live.com"
	omsgBody := "ghwdfushfjhdsjfhjdhfjdhfjdhfjdhfjdh"
	omsgAttach := "D:\\Development\\go\\chrome-native-host-log.txt"
	omsg := oleutil.MustCallMethod(outlook, "createitem", 0).ToIDispatch()
	omsgRec := oleutil.MustCallMethod(omsg, "Recipients").ToIDispatch()
	omsgAtt := oleutil.MustCallMethod(omsg, "Attachments").ToIDispatch()

	oleutil.MustCallMethod(omsgRec, "Add", omsgName).ToIDispatch()
	oleutil.MustPutProperty(omsg, "Subject", omsgSub).ToIDispatch()
	oleutil.MustPutProperty(omsg, "Body", omsgBody).ToIDispatch()

	oleutil.MustCallMethod(omsgAtt, "Add", omsgAttach).ToIDispatch()

	oleutil.MustCallMethod(omsg, "display").ToIDispatch()
	//	oleutil.MustCallMethod(outlook, "Quit")
}
