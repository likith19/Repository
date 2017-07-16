package com.prc.rest.service;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.prc.rest.model.MessageVO;

public class MessageService {

	
	public List<MessageVO> getAllMessages(){
		MessageVO m1 = new MessageVO(1000, "Hello World!!", new Date(),"Likith");
		MessageVO m2 = new MessageVO(1001, "Hello Jersey!!", new Date(),"Likith");
		List<MessageVO> mList = new ArrayList<>();
		mList.add(m1);
		mList.add(m2);
		return mList;
	}
	
}
