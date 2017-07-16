package com.prc.rest;

import java.awt.PageAttributes.MediaType;
import java.util.List;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;

import com.prc.rest.model.MessageVO;
import com.prc.rest.service.MessageService;

@Path("messages")
public class Messages {

	MessageService service = new MessageService();
	
	@GET
	@Produces(javax.ws.rs.core.MediaType.APPLICATION_XML)
	public List<MessageVO> getMessages(){
		return service.getAllMessages();
	}
}
