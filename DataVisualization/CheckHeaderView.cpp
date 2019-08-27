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

/*���Ʊ�ͷlogicalIndex��ʾ��ͷ�ڼ��У�0��ʾ��0�У�Ҳ��������Ҫ����ѡ����У�rect�Ǳ�ͷ��һ�еĴ�С*/

void CheckHeaderView::paintSection(QPainter* painter, const QRect& rect, int logicalIndex) const

{
	painter->save();

	QHeaderView::paintSection(painter, rect, logicalIndex);

	painter->restore();

	if (logicalIndex == 0)//��һ��

	{

		//�����һ�е�λ��

		m_RectHeaderCheckBox.setX(rect.x() + 3);

		m_RectHeaderCheckBox.setY(rect.y() + 10);

		m_RectHeaderCheckBox.setWidth(14);

		m_RectHeaderCheckBox.setHeight(14);

		QStyleOptionButton option;

		QPixmap pix;

		if (m_isOn)

		{//ͼƬ��Ҫ����Դ�����

			pix.load(":/checkImage/image/checkImage/check-sel.png");//ѡ��

			option.state = QStyle::State_On;

		}
		else if (m_RectHeaderCheckBox.contains(m_mousePoint))

		{

			pix.load(":/checkImage/image/checkImage/check-hov.png");//����

			option.state = QStyle::State_HasFocus;

		}
		else
		{

			pix.load(":/checkImage/image/checkImage/check-nor.png");//����

			option.state = QStyle::State_Off;
		}

		style()->drawItemPixmap(painter, rect, Qt::AlignLeft | Qt::AlignVCenter, pix);

		//����2����ѡ��ť,���޷����ñ���ͼƬ

		//

		//option.rect = m_RectHeaderCheckBox;

		//QIcon icon(pix);

		//option.icon = icon;//setStyleSheet("background");

		//style()->drawControl(QStyle::CE_CheckBox, &option, painter);

	}

}



void CheckHeaderView::mousePressEvent(QMouseEvent* event)

{

	//��ͷ�ɿ��ƣ�����һ���У�����ھ��η�Χ��

	if (isEnabled() && logicalIndexAt(event->pos()) == 0 && m_RectHeaderCheckBox.contains(event->pos()))

	{

		m_isOn = !m_isOn;

		updateSection(0);//ˢ�±�ͷ��һ��

		emit headCheckBoxToggled(m_isOn);

	}

	else QHeaderView::mousePressEvent(event);

}

void CheckHeaderView::mouseMoveEvent(QMouseEvent* event)

{

	//��������λ��

	m_mousePoint = event->pos();

	if (m_RectHeaderCheckBox.contains(m_mousePoint))

	{

		updateSection(0);//�ػ��ͷ��ѡ��

	}

	QHeaderView::mouseMoveEvent(event);

}
