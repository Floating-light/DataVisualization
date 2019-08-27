#include "CheckHeaderView.h"



CheckHeaderView::CheckHeaderView(Qt::Orientation orientation,QWidget* parent = nullptr):
	QHeaderView(orientation, parent)
	, m_isOn(false)
{
	m_isOn = false;

	m_mousePoint.setX(100);

	m_mousePoint.setY(100);
}

CheckHeaderView::~CheckHeaderView()
{
}

/*绘制表头logicalIndex表示表头第几列，0表示第0列，也就是我们要画复选框的列，rect是表头第一列的大小*/

void CheckHeaderView::paintSection(QPainter* painter, const QRect& rect, int logicalIndex) const

{
	painter->save();

	QHeaderView::paintSection(painter, rect, logicalIndex);

	painter->restore();

	if (logicalIndex == 0)//第一列

	{

		//保存第一列的位置

		m_RectHeaderCheckBox.setX(rect.x() + 3);

		m_RectHeaderCheckBox.setY(rect.y() + 10);

		m_RectHeaderCheckBox.setWidth(14);

		m_RectHeaderCheckBox.setHeight(14);

		QStyleOptionButton option;

		QPixmap pix;

		if (m_isOn)

		{//图片需要在资源里添加

			pix.load(":/checkImage/image/checkImage/check-sel.png");//选中

			option.state = QStyle::State_On;

		}
		else if (m_RectHeaderCheckBox.contains(m_mousePoint))

		{

			pix.load(":/checkImage/image/checkImage/check-hov.png");//高亮

			option.state = QStyle::State_HasFocus;

		}
		else
		{

			pix.load(":/checkImage/image/checkImage/check-nor.png");//正常

			option.state = QStyle::State_Off;
		}

		style()->drawItemPixmap(painter, rect, Qt::AlignLeft | Qt::AlignVCenter, pix);

		//方法2，画选择按钮,，无法设置背景图片

		//

		//option.rect = m_RectHeaderCheckBox;

		//QIcon icon(pix);

		//option.icon = icon;//setStyleSheet("background");

		//style()->drawControl(QStyle::CE_CheckBox, &option, painter);

	}

}



void CheckHeaderView::mousePressEvent(QMouseEvent* event)

{

	//表头可控制，鼠标第一个列，鼠标在矩形范围内

	if (isEnabled() && logicalIndexAt(event->pos()) == 0 && m_RectHeaderCheckBox.contains(event->pos()))

	{

		m_isOn = !m_isOn;

		updateSection(0);//刷新表头第一列

		emit headCheckBoxToggled(m_isOn);

	}

	else QHeaderView::mousePressEvent(event);

}

void CheckHeaderView::mouseMoveEvent(QMouseEvent* event)

{

	//保存鼠标的位置

	m_mousePoint = event->pos();

	if (m_RectHeaderCheckBox.contains(m_mousePoint))

	{

		updateSection(0);//重绘表头复选框

	}

	QHeaderView::mouseMoveEvent(event);

}
